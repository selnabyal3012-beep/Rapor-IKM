import React, { useState, useEffect } from 'react';
import { ClipboardCheck, Save, CheckCircle2, Compass, Sparkles, FolderDown } from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { P5Score, P5Predicate } from '../../types';
import { StandardActionToolbar } from '../common/StandardActionToolbar';
import { SaveStatusIndicator } from '../common/SaveStatusIndicator';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { useSaveFeedback } from '../../hooks/useSaveFeedback';
import * as XLSX from 'xlsx';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import { buildUniqueP5ScoreId, matchesStudentRecord } from '../../utils/studentSyncHelper';

export const P5ScoreSection: React.FC = () => {
  const appContext = useApp();
  const { db, selectedKelas, selectedSemester, filteredStudent, saveP5Nilai, bulkSaveP5Nilai, showSaveNotification, setActiveMenu } = appContext;
  const { status: saveStatus, triggerSave, lastSavedTime } = useSaveFeedback();
  const [savedNotice, setSavedNotice] = useState(false);

  const projekKelas = db.p5Projek.filter((p) => p.kelas === selectedKelas);
  const [selectedProjekId, setSelectedProjekId] = useState<string>(
    projekKelas[0]?.idProjek || db.p5Projek[0]?.idProjek || 'P001'
  );

  // Pastikan selectedProjekId selalu valid saat kelas atau daftar projek berubah
  useEffect(() => {
    if (projekKelas.length > 0 && !projekKelas.some((p) => p.idProjek === selectedProjekId)) {
      setSelectedProjekId(projekKelas[0].idProjek);
    }
  }, [selectedKelas, db.p5Projek]);

  const currentProjek = projekKelas.find((p) => p.idProjek === selectedProjekId) || projekKelas[0] || db.p5Projek[0];
  const siswaKelas = db.siswa.filter((s) => s.kelas === selectedKelas);
  const displayedSiswa = filteredStudent
    ? siswaKelas.filter((s) => s.idSiswa === filteredStudent.idSiswa)
    : siswaKelas;

  // Pick dimensions to assess (dari db.p5Dimensi atau default 3 dimensi utama)
  const targetDimensiList = React.useMemo(() => {
    const dimMap = new Map<string, { dimensi: string; elemen: string }>();
    (db.p5Dimensi || []).forEach((d) => {
      if (d.dimensi && !dimMap.has(d.dimensi)) {
        dimMap.set(d.dimensi, { dimensi: d.dimensi, elemen: d.elemen || 'Kolaborasi' });
      }
    });
    if (dimMap.size > 0) {
      return Array.from(dimMap.values()).slice(0, 6);
    }
    return [
      { dimensi: 'Gotong Royong', elemen: 'Kolaborasi' },
      { dimensi: 'Mandiri', elemen: 'Regulasi Diri' },
      { dimensi: 'Kreatif', elemen: 'Menghasilkan Karya' }
    ];
  }, [db.p5Dimensi]);

  const defaultDesc = (capaian: P5Predicate, dim: string, nama: string) => {
    if (capaian === 'SB') {
      return `Sangat berkembang dalam dimensi ${dim}, berinisiatif aktif, dan menjadi teladan bagi teman.`;
    } else if (capaian === 'BSH') {
      return `Mampu menunjukkan perkembangan sesuai harapan dalam dimensi ${dim} pada kegiatan projek.`;
    } else if (capaian === 'MB') {
      return `Mulai berkembang dalam dimensi ${dim}, memerlukan sedikit dorongan dari guru dan kawan.`;
    } else {
      return `Belum berkembang dalam dimensi ${dim}, membutuhkan bimbingan dan pendampingan bertahap.`;
    }
  };

  const buildP5ScoresMap = (projId: string): Record<string, P5Score> => {
    const map: Record<string, P5Score> = {};
    siswaKelas.forEach((s) => {
      targetDimensiList.forEach((td) => {
        const key = `${s.idSiswa}_${td.dimensi}`;
        const match = db.p5Nilai.find(
          (n) => n.idProjek === projId && matchesStudentRecord(n, s) && n.dimensi === td.dimensi
        );
        if (match) {
          map[key] = { ...match, idSiswa: s.idSiswa, nisn: s.nisn, namaSiswa: s.namaSiswa, kelas: selectedKelas };
        } else {
          const cap: P5Predicate = 'BSH';
          map[key] = {
            id: buildUniqueP5ScoreId(projId, s.nisn, td.dimensi),
            idProjek: projId,
            idSiswa: s.idSiswa,
            nisn: s.nisn,
            namaSiswa: s.namaSiswa,
            kelas: selectedKelas,
            semester: selectedSemester,
            tahunPelajaran: db.sekolah.tahunPelajaran,
            dimensi: td.dimensi,
            elemen: td.elemen,
            capaian: cap,
            deskripsi: defaultDesc(cap, td.dimensi, s.namaSiswa)
          };
        }
      });
    });
    return map;
  };

  // Local state map: [studentId_dimension] -> P5Score
  const [scores, setScores] = useState<Record<string, P5Score>>(() => buildP5ScoresMap(selectedProjekId));

  useEffect(() => {
    setScores(buildP5ScoresMap(selectedProjekId));
  }, [db.siswa, db.p5Nilai, targetDimensiList, selectedKelas, selectedSemester, selectedProjekId]);

  const handleProjekChange = (projId: string) => {
    setSelectedProjekId(projId);
    setScores(buildP5ScoresMap(projId));
  };

  const getFallbackScore = (idSiswa: string, dim: string): P5Score => {
    const s = siswaKelas.find((x) => x.idSiswa === idSiswa);
    const nisn = s?.nisn || idSiswa;
    const td = targetDimensiList.find((d) => d.dimensi === dim) || { dimensi: dim, elemen: 'Kolaborasi' };
    return {
      id: buildUniqueP5ScoreId(selectedProjekId, nisn, dim),
      idProjek: selectedProjekId,
      idSiswa,
      nisn,
      namaSiswa: s?.namaSiswa || '',
      kelas: selectedKelas,
      semester: selectedSemester,
      tahunPelajaran: db.sekolah.tahunPelajaran,
      dimensi: dim,
      elemen: td.elemen,
      capaian: 'BSH',
      deskripsi: defaultDesc('BSH', dim, s?.namaSiswa || '')
    };
  };

  const updatePredicate = (idSiswa: string, dim: string, cap: P5Predicate) => {
    const key = `${idSiswa}_${dim}`;
    setScores((prev) => {
      const current = prev[key] || getFallbackScore(idSiswa, dim);
      const s = siswaKelas.find((x) => x.idSiswa === idSiswa);
      return {
        ...prev,
        [key]: {
          ...current,
          idSiswa,
          nisn: s?.nisn || current.nisn,
          namaSiswa: s?.namaSiswa || current.namaSiswa,
          capaian: cap,
          deskripsi: defaultDesc(cap, dim, s?.namaSiswa || '')
        }
      };
    });
  };

  const updateDescription = (idSiswa: string, dim: string, desc: string) => {
    const key = `${idSiswa}_${dim}`;
    setScores((prev) => ({
      ...prev,
      [key]: {
        ...(prev[key] || getFallbackScore(idSiswa, dim)),
        deskripsi: desc
      }
    }));
  };

  const handleSaveAll = () => {
    triggerSave(() => {
      const list = Object.values(scores);
      bulkSaveP5Nilai(list);
      setSavedNotice(true);
      setTimeout(() => setSavedNotice(false), 3000);
    });
  };

  const handleClearScores = () => {
    if (confirm('Reset seluruh capaian dimensi P5 menjadi Belum Berkembang (BB)?')) {
      const resetMap: Record<string, P5Score> = {};
      Object.keys(scores).forEach((k) => {
        resetMap[k] = { ...scores[k], capaian: 'BB' };
      });
      setScores(resetMap);
    }
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const data: (string | number)[][] = [
      [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
      [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
      [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
      ['PENILAIAN DIMENSI PROJEK P5'],
      [`Projek: ${currentProjek?.namaProjek || selectedProjekId} • Kelas: ${selectedKelas} • Semester: ${selectedSemester}`],
      [],
      ['NO', 'NISN', 'NAMA SISWA', 'DIMENSI', 'ELEMEN', 'CAPAIAN (BB/MB/BSH/SB)', 'CATATAN PROSES'],
      ['(1)', '(2)', '(3)', '(4)', '(5)', '(6)', '(7)'],
      ...Object.values(scores).map((sc, idx) => [
        idx + 1,
        sc.nisn,
        sc.namaSiswa,
        sc.dimensi,
        sc.elemen,
        sc.capaian,
        sc.deskripsi
      ])
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 6 }, { wch: 18 }, { wch: 30 }, { wch: 20 }, { wch: 20 }, { wch: 26 }, { wch: 55 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Nilai P5');
    XLSX.writeFile(wb, `nilai-p5-${selectedProjekId}-kelas-${selectedKelas}.xlsx`);
  };

  const handleImportFile = async (file: File) => {
    try {
      const result = await parseAndImportTemplate(file, 'template_p5_nilai', appContext);
      if (result.success) {
        showSaveNotification(result.message);
        setTimeout(() => {
          setScores((prev) => {
            const next = { ...prev };
            const currentList = appContext.db.p5Nilai;
            siswaKelas.forEach((s) => {
              targetDimensiList.forEach((td) => {
                const key = `${s.idSiswa}_${td.dimensi}`;
                const found = currentList.find(
                  (p) => matchesStudentRecord(p, s) && p.dimensi === td.dimensi && p.idProjek === selectedProjekId
                );
                if (found) {
                  next[key] = { ...found, idSiswa: s.idSiswa, nisn: s.nisn, namaSiswa: s.namaSiswa };
                }
              });
            });
            return next;
          });
        }, 100);
      } else {
        showSaveNotification(`Gagal impor: ${result.message}`);
      }
    } catch (e: any) {
      console.error(e);
      showSaveNotification('Gagal mengimpor file nilai P5: ' + (e?.message || 'Format tidak sesuai'));
    }
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="p-3 bg-emerald-50 text-emerald-700 rounded-xl border border-emerald-100">
            <ClipboardCheck size={28} />
          </div>
          <div>
            <span className="text-[11px] font-bold tracking-wider uppercase text-emerald-700">
              Penilaian Kokurikuler P5
            </span>
            <h2 className="text-xl font-bold text-slate-800">Input Nilai & Capaian Dimensi P5</h2>
            <p className="text-xs text-slate-500">
              Tabel NILAI_P5: Penilaian 4 skala (BB, MB, BSH, SB) dan catatan proses peserta didik dalam pelaksanaan projek.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          {savedNotice && (
            <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-emerald-900 text-white border border-emerald-400 text-xs font-bold shadow-xs animate-pulse">
              <CheckCircle2 size={14} className="text-emerald-300" />
              <span>Tersimpan</span>
            </span>
          )}
          <button
            onClick={() => setActiveMenu('template_data')}
            className="w-8 h-8 inline-flex items-center justify-center bg-emerald-50 hover:bg-emerald-100 text-emerald-800 rounded-lg transition-colors cursor-pointer border border-emerald-200"
            title="Unduh Template Nilai P5"
            aria-label="Unduh Template Nilai P5"
          >
            <FolderDown size={15} />
          </button>
          <button
            onClick={handleSaveAll}
            disabled={saveStatus === 'saving'}
            title="Simpan Semua Capaian P5"
            aria-label="Simpan Semua Capaian P5"
            className="w-8 h-8 inline-flex items-center justify-center bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg transition-all shadow-xs cursor-pointer disabled:opacity-60 disabled:cursor-not-allowed"
          >
            <Save size={15} />
          </button>
          <SaveStatusIndicator
            status={saveStatus}
            compact
            lastSavedText={lastSavedTime}
          />
        </div>
      </div>

      {/* BARIS TOMBOL AKSI LENGKAP NILAI P5 DI SAMPING / DI ATAS INPUT */}
      <StandardActionToolbar
        title="Aksi Nilai P5"
        badge={`Projek: ${selectedProjekId}`}
        onSave={handleSaveAll}
        saveLabel="Simpan Nilai P5"
        saveStatus={saveStatus}
        onUpdate={handleSaveAll}
        updateLabel="Update Nilai P5"
        onEdit={handleSaveAll}
        editLabel="Edit Nilai P5"
        onCancel={() => {}}
        cancelLabel="Batal"
        onDelete={handleClearScores}
        deleteLabel="Hapus / Reset ke BB"
        onImport={handleImportFile}
        importLabel="Impor Excel (.xlsx)"
        onExport={handleExportExcel}
        exportLabel="Expor Excel (.xlsx)"
        onUpload={handleImportFile}
        uploadLabel="Upload File"
        onClear={handleClearScores}
        clearLabel="Bersihkan File"
        templateLabel="Template P5"
      />

      {/* Project Selector & Legend Banner */}
      <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-4">
        <div className="flex flex-wrap items-center gap-3">
          <div className="flex items-center gap-2">
            <Compass size={18} className="text-emerald-600" />
            <span className="text-xs font-bold text-slate-700">Pilih Projek P5:</span>
            <select
              value={selectedProjekId}
              onChange={(e) => handleProjekChange(e.target.value)}
              className="px-3 py-1.5 text-xs font-bold bg-slate-50 border border-slate-300 rounded-xl text-emerald-800 focus:outline-hidden"
            >
              {projekKelas.map((p) => (
                <option key={p.idProjek} value={p.idProjek}>
                  {p.namaProjek} ({p.tema})
                </option>
              ))}
            </select>
          </div>
          <StudentFilterSelect />
        </div>

        {/* Legend */}
        <div className="flex flex-wrap items-center gap-2 text-[10px] font-bold">
          <span className="px-2 py-0.5 rounded bg-rose-100 text-rose-800">BB: Belum Berkembang</span>
          <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800">MB: Mulai Berkembang</span>
          <span className="px-2 py-0.5 rounded bg-blue-100 text-blue-800">BSH: Berkembang Sesuai Harapan</span>
          <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">SB: Sangat Berkembang</span>
        </div>
      </div>

      {/* Student List Grid */}
      <div className="space-y-4">
        {displayedSiswa.map((s, idx) => (
          <div
            key={s.idSiswa}
            className="bg-white rounded-2xl p-5 border border-slate-200 shadow-xs hover:border-emerald-200 transition-colors"
          >
            <div className="flex items-center gap-2 mb-3 pb-2 border-b border-slate-100">
              <span className="w-6 h-6 rounded-full bg-slate-100 text-slate-600 font-bold text-xs flex items-center justify-center font-mono">
                {idx + 1}
              </span>
              <h3 className="font-bold text-slate-800 text-sm">{s.namaSiswa}</h3>
              <span className="text-xs text-slate-400 font-mono">({s.nisn})</span>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              {targetDimensiList.map((td) => {
                const key = `${s.idSiswa}_${td.dimensi}`;
                const item = scores[key] || {
                  capaian: 'BSH' as P5Predicate,
                  deskripsi: ''
                };

                return (
                  <div key={td.dimensi} className="p-3 rounded-xl bg-slate-50 border border-slate-100 space-y-2">
                    <div className="flex items-center justify-between">
                      <span className="font-bold text-slate-800 text-xs">{td.dimensi}</span>
                      <span className="text-[10px] text-slate-500">{td.elemen}</span>
                    </div>

                    {/* Predicate Selector */}
                    <div className="grid grid-cols-4 gap-1">
                      {(['BB', 'MB', 'BSH', 'SB'] as P5Predicate[]).map((pred) => (
                        <button
                          key={pred}
                          type="button"
                          onClick={() => updatePredicate(s.idSiswa, td.dimensi, pred)}
                          className={`py-1 rounded-lg text-center font-bold text-xs transition-all cursor-pointer ${
                            item.capaian === pred
                              ? pred === 'SB'
                                ? 'bg-emerald-600 text-white shadow-xs'
                                : pred === 'BSH'
                                ? 'bg-blue-600 text-white shadow-xs'
                                : pred === 'MB'
                                ? 'bg-amber-500 text-white shadow-xs'
                                : 'bg-rose-600 text-white shadow-xs'
                              : 'bg-white text-slate-600 border border-slate-200 hover:bg-slate-100'
                          }`}
                        >
                          {pred}
                        </button>
                      ))}
                    </div>

                    {/* Description Text */}
                    <textarea
                      rows={2}
                      value={item.deskripsi}
                      onChange={(e) => updateDescription(s.idSiswa, td.dimensi, e.target.value)}
                      placeholder="Catatan proses capaian..."
                      className="w-full px-2 py-1 text-[11px] border border-slate-200 rounded-lg leading-snug bg-white"
                    />
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};
