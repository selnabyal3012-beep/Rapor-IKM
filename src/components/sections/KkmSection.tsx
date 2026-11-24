import React, { useState, useMemo, useEffect } from 'react';
import {
  Home,
  HelpCircle,
  Save,
  Download,
  Upload,
  Printer,
  FolderDown,
  Zap,
  Sparkles,
  CheckCircle2,
  Calculator,
  X,
  Info
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { Subject, KkmGlobalSettings, KktpRangeItem } from '../../types';
import { DEFAULT_KKM_SETTINGS } from '../../data/defaultData';
import { SaveStatusIndicator } from '../common/SaveStatusIndicator';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { useSaveFeedback } from '../../hooks/useSaveFeedback';
import * as XLSX from 'xlsx';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import {
  buildKktpDescriptionSentence,
  formatStudentCallName,
  resolveKktpRangeForScore
} from '../../utils/tpDescriptionHelper';

const SENI_PILIHAN_DEFS = [
  {
    key: 'seniMusik' as const,
    letter: 'a.',
    kode: 'SENI-MUSIK',
    name: 'Seni Musik',
    badgeBg: 'bg-amber-300 text-slate-950 border-slate-950'
  },
  {
    key: 'seniTari' as const,
    letter: 'b.',
    kode: 'SENI-TARI',
    name: 'Seni Tari',
    badgeBg: 'bg-[#00ff00] text-slate-950 border-slate-950'
  },
  {
    key: 'seniRupa' as const,
    letter: 'c.',
    kode: 'SENI-RUPA',
    name: 'Seni Rupa',
    badgeBg: 'bg-blue-600 text-white border-slate-950'
  },
  {
    key: 'seniTeater' as const,
    letter: 'd.',
    kode: 'SENI-TEATER',
    name: 'Seni Teater',
    badgeBg: 'bg-orange-400 text-slate-950 border-slate-950'
  }
];

const MULOK_LETTERS = ['a.', 'b.', 'c.', 'd.', 'e.'] as const;

type HelpTopic = 'none' | 'deskripsi' | 'mapel' | 'seni' | 'mulok' | 'bobot' | 'kktp';

const HELP_MESSAGES: Record<Exclude<HelpTopic, 'none'>, { title: string; body: string }> = {
  deskripsi: {
    title: 'Opsi Deskripsi Rapor Otomatis',
    body: 'Centang "Deskripsi Rapor Dengan Rentang Predikat" agar kalimat awal deskripsi mengikuti tabel KKTP (Sangat baik / Baik / Cukup / Perlu bimbingan). Centang "Deskripsi Rapor Diawali Nama Panggilan Siswa" agar deskripsi otomatis diawali nama panggilan siswa (misalnya "Manuela menunjukkan penguasaan yang sangat baik dalam ...").'
  },
  mapel: {
    title: 'Daftar Muatan / Mata Pelajaran & Nilai KKM',
    body: 'Isi Nilai KKM (0–100) pada setiap mata pelajaran Wajib, Seni Pilihan, maupun Muatan Lokal. Setelah mengisi semua nilai KKM, klik tombol "Masukan Rumus & Hubungkan ke Semua Mapel" agar KKM dan rumus nilai akhir langsung tersambung ke Input Nilai, Rekap Nilai, dan Rapor.'
  },
  seni: {
    title: 'Mata Pelajaran Seni Pilihan (a–d)',
    body: 'Centang cabang Seni yang diajarkan di satuan pendidikan (Seni Musik, Seni Tari, Seni Rupa, Seni Teater) lalu isi KKM masing-masing. Cabang seni yang dicentang akan otomatis terdaftar pada struktur mata pelajaran dan TP.'
  },
  mulok: {
    title: 'Muatan Lokal (1–5)',
    body: 'Pilih jumlah mata pelajaran Muatan Lokal (1 sampai 5), ketik nama mata pelajaran muatan lokal sekolah (misalnya Bahasa Daerah / Noken Papua), dan isi nilai KKM-nya.'
  },
  bobot: {
    title: 'Rumus Bobot Nilai Rapor',
    body: 'Atur bobot S. TP (Sumatif Tujuan Pembelajaran), S. LM (Sumatif Lingkup Materi), dan S. AS (Sumatif Akhir Semester). Jika baris pertama diubah menjadi FORMATIF, bobot S. TP otomatis menjadi 0 sehingga hanya S. LM dan S. AS yang dihitung ke Nilai Akhir Rapor.'
  },
  kktp: {
    title: 'Kriteria Ketercapaian Tujuan Pembelajaran (KKTP)',
    body: 'Sesuaikan rentang nilai, predikat, dan kalimat awal deskripsi rapor pada 4 tingkat capaian. Kalimat awal ini akan digabung otomatis dengan Tujuan Pembelajaran (TP) tertinggi dan terendah pada menu Input Nilai dan Cetak Rapor.'
  }
};

function isSeniPilihanName(name: string): boolean {
  const n = name.trim().toLowerCase();
  return (
    n === 'seni musik' ||
    n === 'seni tari' ||
    n === 'seni rupa' ||
    n === 'seni teater'
  );
}

function isCombinedSeniName(name: string): boolean {
  const n = name.trim().toLowerCase();
  return n.includes('seni') && n.includes('budaya');
}

export const KkmSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    filteredStudent,
    saveKkmSettingsAndApplyFormula,
    showSaveNotification,
    setActiveMenu
  } = appContext;
  const { status: saveStatus, triggerSave, lastSavedTime } = useSaveFeedback();

  // Merge existing db.kkmSettings with defaults and db.mataPelajaran
  const buildInitialSettings = (): KkmGlobalSettings => {
    const base = db.kkmSettings || DEFAULT_KKM_SETTINGS;
    const existingMulok = db.mataPelajaran
      .filter((m) => m.kelompok === 'Muatan Lokal')
      .map((m) => m.mataPelajaran);
    const mergedMulokNames = [0, 1, 2, 3, 4].map(
      (i) => base.muatanLokalNames?.[i] ?? existingMulok[i] ?? ''
    );
    return {
      deskripsiDenganRentangPredikat: base.deskripsiDenganRentangPredikat ?? true,
      deskripsiDiawaliNamaPanggilan: base.deskripsiDiawaliNamaPanggilan ?? true,
      bobot: {
        namaPenilaianTp: base.bobot?.namaPenilaianTp || 'SUMATIF TUJUAN PEMBELAJARAN',
        bobotTp: base.bobot?.bobotTp ?? 1,
        bobotLm: base.bobot?.bobotLm ?? 1,
        bobotSas: base.bobot?.bobotSas ?? 2
      },
      kktpRanges:
        base.kktpRanges && base.kktpRanges.length === 4
          ? base.kktpRanges.map((r) => ({ ...r }))
          : DEFAULT_KKM_SETTINGS.kktpRanges.map((r) => ({ ...r })),
      seniPilihanAktif: {
        seniMusik: base.seniPilihanAktif?.seniMusik ?? true,
        seniTari: base.seniPilihanAktif?.seniTari ?? true,
        seniRupa: base.seniPilihanAktif?.seniRupa ?? true,
        seniTeater: base.seniPilihanAktif?.seniTeater ?? true
      },
      muatanLokalCount: Math.min(5, Math.max(1, base.muatanLokalCount ?? Math.max(1, existingMulok.length || 5))),
      muatanLokalNames: mergedMulokNames
    };
  };

  const [settings, setSettings] = useState<KkmGlobalSettings>(buildInitialSettings);
  const [activeHelp, setActiveHelp] = useState<HelpTopic>('none');
  const [showSkModal, setShowSkModal] = useState(false);

  // KKM values keyed by subject name (lowercase trimmed) + Seni Pilihan + Muatan Lokal slots
  const [kkmBySubjectKey, setKkmBySubjectKey] = useState<Record<string, number>>(() => {
    const map: Record<string, number> = {};
    db.mataPelajaran.forEach((m) => {
      map[m.mataPelajaran.trim().toLowerCase()] = m.kkm ?? 75;
    });
    SENI_PILIHAN_DEFS.forEach((s) => {
      const key = s.name.toLowerCase();
      if (map[key] === undefined) {
        const combined = db.mataPelajaran.find((m) => isCombinedSeniName(m.mataPelajaran));
        map[key] = combined?.kkm ?? 75;
      }
    });
    return map;
  });

  const [mulokKkmBySlot, setMulokKkmBySlot] = useState<number[]>(() => {
    const existingMulok = db.mataPelajaran.filter((m) => m.kelompok === 'Muatan Lokal');
    return [0, 1, 2, 3, 4].map((i) => existingMulok[i]?.kkm ?? 75);
  });

  // State SK Penetapan KKM (for printing)
  const [skNumber, setSkNumber] = useState(
    `421.2/024/SK-KKM/${db.sekolah.namaSekolah?.replace(/\s+/g, '-').toUpperCase() || 'SDN-BALIL'}/2026`
  );
  const [skDate, setSkDate] = useState('15 Juli 2026');

  // Re-sync when db.kkmSettings or db.mataPelajaran changes externally (e.g. Excel import)
  useEffect(() => {
    setKkmBySubjectKey((prev) => {
      const next = { ...prev };
      db.mataPelajaran.forEach((m) => {
        const k = m.mataPelajaran.trim().toLowerCase();
        if (next[k] === undefined || m.kkm !== undefined) {
          next[k] = m.kkm ?? next[k] ?? 75;
        }
      });
      return next;
    });
  }, [db.mataPelajaran]);

  // Split subjects into structured groups matching the Excel sheet
  const { wajibBeforeSeni, wajibAfterSeni } = useMemo(() => {
    const nonSeniNonMulok = db.mataPelajaran.filter(
      (m) =>
        m.kelompok !== 'Muatan Lokal' &&
        !isSeniPilihanName(m.mataPelajaran) &&
        !isCombinedSeniName(m.mataPelajaran)
    );

    const before: Subject[] = [];
    const after: Subject[] = [];

    nonSeniNonMulok.forEach((m) => {
      const lower = m.mataPelajaran.toLowerCase();
      const isAfterSubject =
        lower.includes('jasmani') ||
        lower.includes('pjok') ||
        lower.includes('olahraga') ||
        lower.includes('inggris') ||
        m.kelompok === 'Pilihan';

      if (isAfterSubject) {
        after.push(m);
      } else {
        before.push(m);
      }
    });

    return { wajibBeforeSeni: before, wajibAfterSeni: after };
  }, [db.mataPelajaran]);

  const combinedSeniSubject = useMemo(
    () => db.mataPelajaran.find((m) => isCombinedSeniName(m.mataPelajaran)),
    [db.mataPelajaran]
  );

  // Build preview / final subject list when applying formula
  const buildMergedMataPelajaran = (): Subject[] => {
    const nextList: Subject[] = [];
    let order = 1;

    // 1. Wajib before Seni
    wajibBeforeSeni.forEach((m) => {
      const kkmVal = kkmBySubjectKey[m.mataPelajaran.trim().toLowerCase()] ?? m.kkm ?? 75;
      nextList.push({
        ...m,
        urutan: order++,
        kkm: kkmVal
      });
    });

    // 2. Preserve combined Seni Budaya if it exists in db so historical TPs/scores keep working, synced with average of active Seni Pilihan
    const activeSeniDefs = SENI_PILIHAN_DEFS.filter(
      (s) => settings.seniPilihanAktif[s.key]
    );
    if (combinedSeniSubject) {
      const avgSeniKkm =
        activeSeniDefs.length > 0
          ? Math.round(
              activeSeniDefs.reduce(
                (acc, s) => acc + (kkmBySubjectKey[s.name.toLowerCase()] ?? 75),
                0
              ) / activeSeniDefs.length
            )
          : combinedSeniSubject.kkm ?? 75;
      nextList.push({
        ...combinedSeniSubject,
        urutan: order++,
        kkm: avgSeniKkm
      });
    }

    // 3. Checked Seni Pilihan subjects (Seni Musik, Seni Tari, Seni Rupa, Seni Teater)
    activeSeniDefs.forEach((seni) => {
      const existing = db.mataPelajaran.find(
        (m) => m.mataPelajaran.trim().toLowerCase() === seni.name.toLowerCase()
      );
      const kkmVal = kkmBySubjectKey[seni.name.toLowerCase()] ?? existing?.kkm ?? 75;
      nextList.push({
        kode: existing?.kode || seni.kode,
        mataPelajaran: seni.name,
        kelompok: existing?.kelompok || 'Pilihan',
        urutan: order++,
        kkm: kkmVal,
        kompleksitas: existing?.kompleksitas ?? kkmVal,
        dayaDukung: existing?.dayaDukung ?? kkmVal,
        intake: existing?.intake ?? kkmVal
      });
    });

    // 4. Wajib after Seni (PJOK, Bahasa Inggris, dll.)
    wajibAfterSeni.forEach((m) => {
      const kkmVal = kkmBySubjectKey[m.mataPelajaran.trim().toLowerCase()] ?? m.kkm ?? 75;
      nextList.push({
        ...m,
        urutan: order++,
        kkm: kkmVal
      });
    });

    // 5. Muatan Lokal active slots (1..muatanLokalCount) with non-empty names
    const existingMulokList = db.mataPelajaran.filter((m) => m.kelompok === 'Muatan Lokal');
    for (let i = 0; i < settings.muatanLokalCount; i++) {
      const rawName = (settings.muatanLokalNames[i] || '').trim();
      if (!rawName) continue;
      const existing =
        existingMulokList.find(
          (m) => m.mataPelajaran.trim().toLowerCase() === rawName.toLowerCase()
        ) || existingMulokList[i];
      const kkmVal = mulokKkmBySlot[i] ?? existing?.kkm ?? 75;
      nextList.push({
        kode: existing?.kode || `MULOK${i + 1}`,
        mataPelajaran: rawName,
        kelompok: 'Muatan Lokal',
        urutan: order++,
        kkm: kkmVal,
        kompleksitas: existing?.kompleksitas ?? kkmVal,
        dayaDukung: existing?.dayaDukung ?? kkmVal,
        intake: existing?.intake ?? kkmVal
      });
    }

    return nextList;
  };

  // Primary Action: Save KKM all subjects + Bobot + KKTP + Deskripsi options & connect to all subjects & reports
  const handleConnectFormulaToAllSubjects = () => {
    const mergedMapels = buildMergedMataPelajaran();
    triggerSave(async () => {
      saveKkmSettingsAndApplyFormula(settings, mergedMapels);
    });
  };

  const handleSubjectKkmChange = (subjectName: string, rawVal: string) => {
    const num = Math.min(100, Math.max(0, Number(rawVal) || 0));
    setKkmBySubjectKey((prev) => ({
      ...prev,
      [subjectName.trim().toLowerCase()]: num
    }));
  };

  const handleApplyUniformKkm = (uniformValue: number) => {
    setKkmBySubjectKey((prev) => {
      const next = { ...prev };
      Object.keys(next).forEach((k) => {
        next[k] = uniformValue;
      });
      wajibBeforeSeni.forEach((m) => {
        next[m.mataPelajaran.trim().toLowerCase()] = uniformValue;
      });
      wajibAfterSeni.forEach((m) => {
        next[m.mataPelajaran.trim().toLowerCase()] = uniformValue;
      });
      SENI_PILIHAN_DEFS.forEach((s) => {
        next[s.name.toLowerCase()] = uniformValue;
      });
      return next;
    });
    setMulokKkmBySlot([uniformValue, uniformValue, uniformValue, uniformValue, uniformValue]);
    showSaveNotification(
      `Nilai KKM seragam ${uniformValue} diisi pada semua mapel. Klik "Masukan Rumus & Hubungkan ke Semua Mapel" untuk menyambungkan.`
    );
  };

  // Bobot Handlers
  const handleChangeNamaPenilaianTp = (mode: 'SUMATIF TUJUAN PEMBELAJARAN' | 'FORMATIF') => {
    setSettings((prev) => ({
      ...prev,
      bobot: {
        ...prev.bobot,
        namaPenilaianTp: mode,
        bobotTp: mode === 'FORMATIF' ? 0 : prev.bobot.bobotTp === 0 ? 1 : prev.bobot.bobotTp
      }
    }));
  };

  const handleChangeBobot = (field: 'bobotTp' | 'bobotLm' | 'bobotSas', rawVal: string) => {
    const num = Math.min(10, Math.max(0, Number(rawVal) || 0));
    setSettings((prev) => ({
      ...prev,
      bobot: {
        ...prev.bobot,
        [field]: num
      }
    }));
  };

  // KKTP Row Handler
  const handleUpdateKktpRow = (
    idx: number,
    field: keyof Pick<KktpRangeItem, 'min' | 'max' | 'labelPredikat' | 'kalimatAwal'>,
    value: string
  ) => {
    setSettings((prev) => {
      const nextRanges = prev.kktpRanges.map((item, rIdx) => {
        if (rIdx !== idx) return item;
        if (field === 'min' || field === 'max') {
          const num = Math.min(100, Math.max(0, Number(value) || 0));
          return { ...item, [field]: num };
        }
        return { ...item, [field]: value };
      });
      return { ...prev, kktpRanges: nextRanges };
    });
  };

  // Live Formula Display
  const activeBobotTp =
    settings.bobot.namaPenilaianTp === 'FORMATIF' ? 0 : settings.bobot.bobotTp;
  const activeBobotLm = settings.bobot.bobotLm;
  const activeBobotSas = settings.bobot.bobotSas;
  const totalBobot = activeBobotTp + activeBobotLm + activeBobotSas;

  const liveFormulaText =
    activeBobotTp > 0
      ? `NA RAPOR = (${activeBobotTp}*S. TP + ${activeBobotLm}*S. LM + ${activeBobotSas}*S. AS)/${totalBobot || 1}`
      : `NA RAPOR = (${activeBobotLm}*S. LM + ${activeBobotSas}*S. AS)/${totalBobot || 1}`;

  // Sample calculation preview for teacher verification
  const previewResult = useMemo(() => {
    const sampleTp = 82;
    const sampleLm = 84;
    const sampleSas = 80;
    const weighted =
      totalBobot > 0
        ? Math.round(
            (activeBobotTp * sampleTp + activeBobotLm * sampleLm + activeBobotSas * sampleSas) /
              totalBobot
          )
        : 0;
    const matchedKktp = resolveKktpRangeForScore(weighted, settings.kktpRanges);
    const callName = formatStudentCallName(
      filteredStudent?.namaSiswa || db.siswa[0]?.namaSiswa || 'Manuela',
      filteredStudent?.bukuInduk?.namaPanggilan || db.siswa[0]?.bukuInduk?.namaPanggilan
    );
    const sampleSentence = buildKktpDescriptionSentence(
      weighted,
      'memahami materi dan menyelesaikan latihan soal dengan tepat',
      settings,
      callName
    );
    return {
      sampleTp,
      sampleLm,
      sampleSas,
      weighted,
      matchedKktp,
      callName,
      sampleSentence
    };
  }, [activeBobotTp, activeBobotLm, activeBobotSas, totalBobot, settings, filteredStudent, db.siswa]);

  // Average KKM summary
  const previewMapels = useMemo(() => buildMergedMataPelajaran(), [
    wajibBeforeSeni,
    wajibAfterSeni,
    combinedSeniSubject,
    settings.seniPilihanAktif,
    settings.muatanLokalCount,
    settings.muatanLokalNames,
    kkmBySubjectKey,
    mulokKkmBySlot
  ]);

  const avgKkm = useMemo(() => {
    if (previewMapels.length === 0) return 75;
    const sum = previewMapels.reduce((acc, m) => acc + (m.kkm ?? 75), 0);
    return Math.round(sum / previewMapels.length);
  }, [previewMapels]);

  // Excel Export & Import
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const rowsData: (string | number)[][] = [
      [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
      [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
      [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
      ['MATA PELAJARAN, NILAI KKM, BOBOT NILAI RAPOR & KKTP'],
      [`Tahun Ajaran: ${db.sekolah.tahunPelajaran} • Rumus: ${liveFormulaText}`],
      [],
      ['NO', 'KODE', 'MATA PELAJARAN', 'KELOMPOK', 'NILAI KKM'],
      ...previewMapels.map((m, idx) => [
        idx + 1,
        m.kode || `M0${idx + 1}`,
        m.mataPelajaran,
        m.kelompok,
        m.kkm ?? 75
      ]),
      [],
      ['BOBOT NILAI RAPOR'],
      ['Komponen', 'Singkatan', 'Bobot'],
      [settings.bobot.namaPenilaianTp, settings.bobot.namaPenilaianTp === 'FORMATIF' ? 'FORMATIF' : 'S. TP', activeBobotTp],
      ['SUMATIF LINGKUP MATERI', 'S. LM', activeBobotLm],
      ['SUMATIF AKHIR SEMESTER', 'S. AS', activeBobotSas],
      [],
      ['KRITERIA KETERCAPAIAN TUJUAN PEMBELAJARAN (KKTP)'],
      ['Rentang Min', 'Rentang Max', 'Predikat', 'Kalimat Awal Deskripsi Rapor'],
      ...settings.kktpRanges.map((r) => [r.min, r.max, r.labelPredikat, r.kalimatAwal])
    ];

    const ws = XLSX.utils.aoa_to_sheet(rowsData);
    ws['!cols'] = [{ wch: 8 }, { wch: 14 }, { wch: 42 }, { wch: 18 }, { wch: 48 }];
    XLSX.utils.book_append_sheet(wb, ws, 'KKM & KKTP');
    XLSX.writeFile(
      wb,
      `kkm-bobot-kktp-${db.sekolah.namaSekolah?.toLowerCase().replace(/\s+/g, '-') || 'sdn-balil'}.xlsx`
    );
    showSaveNotification('File KKM, Bobot Nilai Rapor & KKTP (.xlsx) berhasil diunduh!');
  };

  const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = '';
    try {
      const result = await parseAndImportTemplate(file, 'template_kkm', appContext);
      if (result.success) {
        showSaveNotification(result.message);
      } else {
        showSaveNotification(`Gagal impor: ${result.message}`);
      }
    } catch (err: any) {
      showSaveNotification('Gagal membaca file Excel KKM: ' + (err?.message || 'Format berkas tidak sesuai'));
    }
  };

  const seniRowCount = SENI_PILIHAN_DEFS.length;
  const mulokSlots = Array.from({ length: settings.muatanLokalCount }, (_, i) => i);

  return (
    <div className="space-y-4">
      {/* MAIN EXCEL-STYLE WORKBOOK CANVAS */}
      <div className="print:hidden bg-[#00ff00] border-2 border-slate-950 rounded-xl p-2 sm:p-3 shadow-md space-y-2.5 text-slate-950">
        {/* TOP HEADER ROW */}
        <div className="grid grid-cols-1 xl:grid-cols-[minmax(320px,1.05fr)_minmax(270px,0.85fr)_minmax(360px,1.35fr)] gap-2.5 items-stretch">
          {/* Left Black Bar: Home + MATA PELAJARAN + Nama Sekolah */}
          <div className="bg-black text-white border-2 border-slate-950 px-3 py-2 flex items-center gap-3">
            <button
              type="button"
              onClick={() => setActiveMenu('dashboard')}
              title="Kembali ke Dashboard / Home"
              aria-label="Kembali ke Dashboard"
              className="w-10 h-10 rounded-full bg-sky-500 hover:bg-sky-400 text-slate-950 flex items-center justify-center border-2 border-sky-200 shadow-xs shrink-0 cursor-pointer transition-colors"
            >
              <Home size={20} strokeWidth={2.4} />
            </button>
            <div className="flex-1 text-center leading-tight">
              <div className="text-base sm:text-lg font-black tracking-wide uppercase font-serif">
                MATA PELAJARAN
              </div>
              <div className="text-sm sm:text-base font-extrabold tracking-wide uppercase font-serif text-slate-100">
                {db.sekolah.namaSekolah || 'SD NEGERI BALIL'}
              </div>
            </div>
          </div>

          {/* Center Description Checkboxes + Purple Help Icon */}
          <div className="flex items-stretch gap-1.5">
            <div className="flex-1 flex flex-col gap-1 justify-center">
              <label className="flex items-center gap-2 bg-[#ff0000] text-white border-2 border-black px-2.5 py-1 text-xs font-bold cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={settings.deskripsiDenganRentangPredikat}
                  onChange={(e) =>
                    setSettings((prev) => ({
                      ...prev,
                      deskripsiDenganRentangPredikat: e.target.checked
                    }))
                  }
                  className="w-4 h-4 accent-white cursor-pointer"
                />
                <span>Deskripsi Rapor Dengan Rentang Predikat</span>
              </label>

              <label className="flex items-center gap-2 bg-white text-slate-950 border-2 border-black px-2.5 py-1 text-xs font-bold cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={settings.deskripsiDiawaliNamaPanggilan}
                  onChange={(e) =>
                    setSettings((prev) => ({
                      ...prev,
                      deskripsiDiawaliNamaPanggilan: e.target.checked
                    }))
                  }
                  className="w-4 h-4 accent-slate-900 cursor-pointer"
                />
                <span>Deskripsi Rapor Diawali Nama Panggilan Siswa</span>
              </label>
            </div>

            <button
              type="button"
              onClick={() => setActiveHelp(activeHelp === 'deskripsi' ? 'none' : 'deskripsi')}
              title="Informasi Opsi Deskripsi Rapor"
              aria-label="Informasi Opsi Deskripsi Rapor"
              className="w-9 bg-black hover:bg-slate-900 text-purple-400 border-2 border-slate-950 flex items-center justify-center cursor-pointer shrink-0"
            >
              <HelpCircle size={20} />
            </button>
          </div>

          {/* Right Heading + Action Buttons */}
          <div className="flex flex-col justify-between gap-1.5">
            <div className="text-center font-serif font-black text-sm sm:text-base tracking-wide text-slate-950 uppercase leading-tight">
              KRITERIA KETERCAPAIAN TUJUAN PEMBELAJARAN (KKTP)
            </div>
            <div className="flex flex-wrap items-center justify-between gap-1.5 bg-white/95 border-2 border-slate-950 px-2 py-1">
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={handleExportExcel}
                  title="Ekspor KKM, Bobot & KKTP ke Excel (.xlsx)"
                  className="w-7 h-7 inline-flex items-center justify-center bg-emerald-600 hover:bg-emerald-700 text-white rounded-sm cursor-pointer"
                >
                  <Download size={14} />
                </button>
                <label
                  title="Impor KKM dari Excel (.xlsx)"
                  className="w-7 h-7 inline-flex items-center justify-center bg-blue-600 hover:bg-blue-700 text-white rounded-sm cursor-pointer"
                >
                  <Upload size={14} />
                  <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleImportExcel}
                    className="hidden"
                  />
                </label>
                <button
                  type="button"
                  onClick={() => setShowSkModal((prev) => !prev)}
                  title="Atur & Cetak SK Penetapan KKM"
                  className="w-7 h-7 inline-flex items-center justify-center bg-slate-800 hover:bg-slate-900 text-white rounded-sm cursor-pointer"
                >
                  <Printer size={14} />
                </button>
                <button
                  type="button"
                  onClick={() => setActiveMenu('template_data')}
                  title="Buka Template Data"
                  className="w-7 h-7 inline-flex items-center justify-center bg-indigo-600 hover:bg-indigo-700 text-white rounded-sm cursor-pointer"
                >
                  <FolderDown size={14} />
                </button>
              </div>

              <div className="flex items-center gap-1.5">
                <SaveStatusIndicator status={saveStatus} compact lastSavedText={lastSavedTime} />
                <button
                  type="button"
                  onClick={handleConnectFormulaToAllSubjects}
                  disabled={saveStatus === 'saving'}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-950 hover:bg-blue-900 text-amber-300 font-black text-[11px] rounded-sm border border-amber-400 shadow-xs cursor-pointer transition-colors disabled:opacity-60"
                >
                  <Zap size={13} className="fill-amber-300" />
                  <span>Masukan Rumus &amp; Hubungkan ke Semua Mapel</span>
                </button>
              </div>
            </div>
          </div>
        </div>

        {/* Help Callout Banner (toggleable via ? icons) */}
        {activeHelp !== 'none' && (
          <div className="bg-white border-2 border-slate-950 px-3 py-2 flex items-start justify-between gap-3 text-xs shadow-sm">
            <div className="flex items-start gap-2">
              <Info size={16} className="text-purple-700 shrink-0 mt-0.5" />
              <div>
                <div className="font-black text-slate-950">{HELP_MESSAGES[activeHelp].title}</div>
                <div className="text-slate-700 leading-snug mt-0.5">
                  {HELP_MESSAGES[activeHelp].body}
                </div>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setActiveHelp('none')}
              className="p-1 hover:bg-slate-100 rounded-sm text-slate-600 cursor-pointer"
              aria-label="Tutup Bantuan"
            >
              <X size={14} />
            </button>
          </div>
        )}

        {/* 3-COLUMN EXCEL WORKBOOK GRID */}
        <div className="grid grid-cols-1 xl:grid-cols-[minmax(350px,1.08fr)_minmax(285px,0.88fr)_minmax(390px,1.34fr)] gap-2.5 items-start">
          {/* COLUMN 1: MUATAN/MATA PELAJARAN + KKM */}
          <div className="space-y-2">
            <div className="bg-white border-2 border-red-600 overflow-x-auto">
              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr className="bg-black text-white font-serif">
                    <th className="border border-red-600 px-2 py-2.5 text-center w-10 font-black">
                      NO.
                    </th>
                    <th className="border border-red-600 px-2.5 py-2.5 text-center font-black">
                      <div className="inline-flex items-center justify-center gap-2">
                        <span>MUATAN/MATA PELAJARAN</span>
                        <button
                          type="button"
                          onClick={() => setActiveHelp(activeHelp === 'mapel' ? 'none' : 'mapel')}
                          title="Bantuan Pengisian Mapel & KKM"
                          className="text-purple-400 hover:text-purple-300 cursor-pointer"
                        >
                          <HelpCircle size={15} />
                        </button>
                      </div>
                    </th>
                    <th className="border border-red-600 px-2 py-2.5 text-center w-16 font-black bg-slate-950 text-amber-300">
                      KKM
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {/* Wajib Before Seni (1..4 + IPAS if active) */}
                  {wajibBeforeSeni.map((m, idx) => {
                    const kkmVal = kkmBySubjectKey[m.mataPelajaran.trim().toLowerCase()] ?? m.kkm ?? 75;
                    return (
                      <tr key={m.kode || m.mataPelajaran} className="bg-white hover:bg-amber-50/40">
                        <td className="border border-red-600 px-1.5 py-1.5 text-center font-bold tabular-nums">
                          {idx + 1}
                        </td>
                        <td className="border border-red-600 px-2.5 py-1.5 font-serif font-bold text-slate-950">
                          {m.mataPelajaran}
                        </td>
                        <td className="border border-red-600 p-0.5 text-center bg-amber-50/60">
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={kkmVal}
                            onChange={(e) => handleSubjectKkmChange(m.mataPelajaran, e.target.value)}
                            aria-label={`KKM ${m.mataPelajaran}`}
                            className="w-14 px-1 py-1 text-center font-mono tabular-nums font-black text-slate-950 bg-white border border-slate-300 focus:border-blue-600 focus:outline-hidden rounded-xs"
                          />
                        </td>
                      </tr>
                    );
                  })}

                  {/* Seni Pilihan Bar */}
                  <tr className="bg-cyan-300">
                    <td colSpan={3} className="border border-red-600 px-2 py-1.5">
                      <div className="flex flex-wrap items-center gap-1.5">
                        <span className="font-serif font-bold text-slate-950 text-xs mr-0.5">
                          Seni Pilihan
                        </span>
                        {SENI_PILIHAN_DEFS.map((seni) => {
                          const checked = settings.seniPilihanAktif[seni.key];
                          return (
                            <label
                              key={seni.key}
                              className={`inline-flex items-center gap-1 px-1.5 py-0.5 text-[10px] font-bold border cursor-pointer select-none ${seni.badgeBg}`}
                            >
                              <input
                                type="checkbox"
                                checked={checked}
                                onChange={(e) =>
                                  setSettings((prev) => ({
                                    ...prev,
                                    seniPilihanAktif: {
                                      ...prev.seniPilihanAktif,
                                      [seni.key]: e.target.checked
                                    }
                                  }))
                                }
                                className="w-3 h-3 cursor-pointer"
                              />
                              <span>{seni.name}</span>
                            </label>
                          );
                        })}
                        <button
                          type="button"
                          onClick={() => setActiveHelp(activeHelp === 'seni' ? 'none' : 'seni')}
                          title="Bantuan Seni Pilihan"
                          className="ml-auto text-blue-800 hover:text-blue-950 cursor-pointer"
                        >
                          <HelpCircle size={15} />
                        </button>
                      </div>
                    </td>
                  </tr>

                  {/* Seni Pilihan Sub-rows a..d */}
                  {SENI_PILIHAN_DEFS.map((seni, sIdx) => {
                    const checked = settings.seniPilihanAktif[seni.key];
                    const kkmVal = kkmBySubjectKey[seni.name.toLowerCase()] ?? 75;
                    const seniRowNum = wajibBeforeSeni.length + 1;
                    return (
                      <tr
                        key={seni.key}
                        className={checked ? 'bg-[#e6def4]' : 'bg-slate-200/75 opacity-60'}
                      >
                        {sIdx === 0 && (
                          <td
                            rowSpan={seniRowCount}
                            className="border border-red-600 px-1.5 py-1.5 text-center font-bold bg-white tabular-nums align-middle"
                          >
                            {seniRowNum}
                          </td>
                        )}
                        <td className="border border-red-600 px-2 py-1.5 font-serif font-bold text-slate-950">
                          <div className="flex items-center gap-2">
                            <span className="w-5 inline-block text-center font-bold">{seni.letter}</span>
                            <span>{seni.name}</span>
                          </div>
                        </td>
                        <td className="border border-red-600 p-0.5 text-center bg-white/80">
                          <input
                            type="number"
                            min={0}
                            max={100}
                            disabled={!checked}
                            value={kkmVal}
                            onChange={(e) => handleSubjectKkmChange(seni.name, e.target.value)}
                            aria-label={`KKM ${seni.name}`}
                            className="w-14 px-1 py-1 text-center font-mono tabular-nums font-black text-slate-950 bg-white border border-slate-300 focus:border-blue-600 focus:outline-hidden rounded-xs disabled:bg-slate-100"
                          />
                        </td>
                      </tr>
                    );
                  })}

                  {/* Wajib After Seni (PJOK, Bahasa Inggris, dll.) */}
                  {wajibAfterSeni.map((m, idx) => {
                    const rowNumber = wajibBeforeSeni.length + 2 + idx;
                    const kkmVal = kkmBySubjectKey[m.mataPelajaran.trim().toLowerCase()] ?? m.kkm ?? 75;
                    return (
                      <tr key={m.kode || m.mataPelajaran} className="bg-white hover:bg-amber-50/40">
                        <td className="border border-red-600 px-1.5 py-1.5 text-center font-bold tabular-nums">
                          {rowNumber}
                        </td>
                        <td className="border border-red-600 px-2.5 py-1.5 font-serif font-bold text-slate-950">
                          {m.mataPelajaran}
                        </td>
                        <td className="border border-red-600 p-0.5 text-center bg-amber-50/60">
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={kkmVal}
                            onChange={(e) => handleSubjectKkmChange(m.mataPelajaran, e.target.value)}
                            aria-label={`KKM ${m.mataPelajaran}`}
                            className="w-14 px-1 py-1 text-center font-mono tabular-nums font-black text-slate-950 bg-white border border-slate-300 focus:border-blue-600 focus:outline-hidden rounded-xs"
                          />
                        </td>
                      </tr>
                    );
                  })}

                  {/* Muatan Lokal Cyan Header Bar */}
                  <tr className="bg-cyan-300">
                    <td colSpan={3} className="border border-red-600 px-2 py-1.5">
                      <div className="flex items-center gap-2">
                        <span className="font-serif font-bold text-slate-950 text-xs">
                          Muatan Lokal
                        </span>
                        <select
                          value={settings.muatanLokalCount}
                          onChange={(e) =>
                            setSettings((prev) => ({
                              ...prev,
                              muatanLokalCount: Number(e.target.value)
                            }))
                          }
                          aria-label="Jumlah Muatan Lokal"
                          className="px-2 py-0.5 bg-white border border-slate-950 text-xs font-bold text-slate-950 cursor-pointer"
                        >
                          {[1, 2, 3, 4, 5].map((n) => (
                            <option key={n} value={n}>
                              {n}
                            </option>
                          ))}
                        </select>
                        <button
                          type="button"
                          onClick={() => setActiveHelp(activeHelp === 'mulok' ? 'none' : 'mulok')}
                          title="Bantuan Muatan Lokal"
                          className="text-blue-800 hover:text-blue-950 cursor-pointer"
                        >
                          <HelpCircle size={15} />
                        </button>
                      </div>
                    </td>
                  </tr>

                  {/* Muatan Lokal Sub-rows a..e */}
                  {mulokSlots.map((slotIdx) => {
                    const mulokRowNumber = wajibBeforeSeni.length + wajibAfterSeni.length + 2;
                    return (
                      <tr key={`mulok-${slotIdx}`} className="bg-white">
                        {slotIdx === 0 && (
                          <td
                            rowSpan={mulokSlots.length}
                            className="border border-red-600 px-1.5 py-1.5 text-center font-bold tabular-nums align-middle"
                          >
                            {mulokRowNumber}
                          </td>
                        )}
                        <td className="border border-red-600 px-2 py-1">
                          <div className="flex items-center gap-1.5">
                            <span className="w-5 inline-block text-center font-bold font-serif shrink-0">
                              {MULOK_LETTERS[slotIdx]}
                            </span>
                            <input
                              type="text"
                              value={settings.muatanLokalNames[slotIdx] || ''}
                              onChange={(e) => {
                                const val = e.target.value;
                                setSettings((prev) => {
                                  const nextNames = [...prev.muatanLokalNames];
                                  nextNames[slotIdx] = val;
                                  return { ...prev, muatanLokalNames: nextNames };
                                });
                              }}
                              placeholder={`Muatan Lokal ${slotIdx + 1} (opsional)`}
                              className="w-full px-1.5 py-0.5 text-xs font-serif font-bold text-slate-950 bg-transparent border border-transparent focus:border-blue-500 focus:bg-white focus:outline-hidden rounded-xs"
                            />
                          </div>
                        </td>
                        <td className="border border-red-600 p-0.5 text-center bg-amber-50/60">
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={mulokKkmBySlot[slotIdx] ?? 75}
                            onChange={(e) => {
                              const num = Math.min(100, Math.max(0, Number(e.target.value) || 0));
                              setMulokKkmBySlot((prev) => {
                                const copy = [...prev];
                                copy[slotIdx] = num;
                                return copy;
                              });
                            }}
                            aria-label={`KKM Muatan Lokal ${slotIdx + 1}`}
                            className="w-14 px-1 py-1 text-center font-mono tabular-nums font-black text-slate-950 bg-white border border-slate-300 focus:border-blue-600 focus:outline-hidden rounded-xs"
                          />
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Uniform KKM Quick Bar + Connect Action */}
            <div className="bg-white border-2 border-red-600 p-2.5 space-y-2">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-1.5 text-[11px] font-extrabold text-slate-800">
                  <Sparkles size={14} className="text-amber-500" />
                  <span>Isi KKM Seragam Cepat:</span>
                </div>
                <div className="flex flex-wrap items-center gap-1">
                  {[70, 72, 75, 78, 80].map((val) => (
                    <button
                      key={val}
                      type="button"
                      onClick={() => handleApplyUniformKkm(val)}
                      className="px-2 py-0.5 text-[11px] font-extrabold bg-emerald-50 hover:bg-emerald-100 text-emerald-900 border border-emerald-400 rounded-xs cursor-pointer"
                    >
                      {val}
                    </button>
                  ))}
                  <span className="ml-1 px-2 py-0.5 text-[10px] font-mono font-bold bg-slate-100 border border-slate-300 text-slate-800">
                    Rata-rata: {avgKkm}
                  </span>
                </div>
              </div>

              <button
                type="button"
                onClick={handleConnectFormulaToAllSubjects}
                disabled={saveStatus === 'saving'}
                className="w-full py-2 px-3 bg-gradient-to-r from-blue-700 to-indigo-800 hover:from-blue-600 hover:to-indigo-700 text-white font-black text-xs rounded-sm border border-slate-950 shadow-xs flex items-center justify-center gap-2 cursor-pointer transition-all disabled:opacity-60"
              >
                <CheckCircle2 size={15} className="text-amber-300" />
                <span>Masukan Rumus &amp; Hubungkan ke Semua Mapel ({previewMapels.length} Mapel)</span>
              </button>
            </div>
          </div>

          {/* COLUMN 2: BOBOT NILAI RAPOR + LIVE FORMULA + INFORMASI */}
          <div className="space-y-2.5">
            <div className="bg-white border-2 border-red-600 overflow-hidden">
              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr className="bg-black text-white font-serif">
                    <th colSpan={3} className="border border-red-600 px-3 py-2.5 text-center font-black">
                      <div className="inline-flex items-center justify-center gap-2">
                        <span className="text-sm">BOBOT NILAI RAPOR</span>
                        <button
                          type="button"
                          onClick={() => setActiveHelp(activeHelp === 'bobot' ? 'none' : 'bobot')}
                          title="Bantuan Bobot Nilai Rapor"
                          className="text-purple-400 hover:text-purple-300 cursor-pointer"
                        >
                          <HelpCircle size={15} />
                        </button>
                      </div>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {/* Row 1: SUMATIF TUJUAN PEMBELAJARAN / FORMATIF (Purple) */}
                  <tr className="bg-[#d9c2e9] text-slate-950 font-serif font-extrabold">
                    <td className="border border-red-600 p-1.5 text-center">
                      <select
                        value={settings.bobot.namaPenilaianTp}
                        onChange={(e) =>
                          handleChangeNamaPenilaianTp(
                            e.target.value as 'SUMATIF TUJUAN PEMBELAJARAN' | 'FORMATIF'
                          )
                        }
                        aria-label="Jenis Penilaian Tujuan Pembelajaran"
                        className="w-full bg-white/80 hover:bg-white border border-slate-600 px-1.5 py-1 text-[11px] font-serif font-black text-center text-slate-950 cursor-pointer focus:outline-hidden"
                      >
                        <option value="SUMATIF TUJUAN PEMBELAJARAN">SUMATIF TUJUAN PEMBELAJARAN</option>
                        <option value="FORMATIF">FORMATIF</option>
                      </select>
                    </td>
                    <td className="border border-red-600 px-2 py-2 text-center w-20 font-black">
                      {settings.bobot.namaPenilaianTp === 'FORMATIF' ? 'FORMATIF' : 'S. TP'}
                    </td>
                    <td className="border border-red-600 p-1 text-center w-14">
                      <input
                        type="number"
                        min={0}
                        max={10}
                        value={activeBobotTp}
                        onChange={(e) => handleChangeBobot('bobotTp', e.target.value)}
                        aria-label="Bobot S. TP"
                        className="w-11 py-1 text-center font-mono tabular-nums font-black text-base text-slate-950 bg-white/85 border border-slate-500 focus:border-blue-700 focus:outline-hidden"
                      />
                    </td>
                  </tr>

                  {/* Row 2: SUMATIF LINGKUP MATERI (Yellow) */}
                  <tr className="bg-[#ffff00] text-slate-950 font-serif font-extrabold">
                    <td className="border border-red-600 px-2 py-2.5 text-center text-[11px]">
                      SUMATIF LINGKUP MATERI
                    </td>
                    <td className="border border-red-600 px-2 py-2.5 text-center w-20 font-black">
                      S. LM
                    </td>
                    <td className="border border-red-600 p-1 text-center w-14">
                      <input
                        type="number"
                        min={0}
                        max={10}
                        value={activeBobotLm}
                        onChange={(e) => handleChangeBobot('bobotLm', e.target.value)}
                        aria-label="Bobot S. LM"
                        className="w-11 py-1 text-center font-mono tabular-nums font-black text-base text-slate-950 bg-white/85 border border-slate-500 focus:border-blue-700 focus:outline-hidden"
                      />
                    </td>
                  </tr>

                  {/* Row 3: SUMATIF AKHIR SEMESTER (Cyan) */}
                  <tr className="bg-[#00ffff] text-slate-950 font-serif font-extrabold">
                    <td className="border border-red-600 px-2 py-2.5 text-center text-[11px]">
                      SUMATIF AKHIR SEMESTER
                    </td>
                    <td className="border border-red-600 px-2 py-2.5 text-center w-20 font-black">
                      S. AS
                    </td>
                    <td className="border border-red-600 p-1 text-center w-14">
                      <input
                        type="number"
                        min={0}
                        max={10}
                        value={activeBobotSas}
                        onChange={(e) => handleChangeBobot('bobotSas', e.target.value)}
                        aria-label="Bobot S. AS"
                        className="w-11 py-1 text-center font-mono tabular-nums font-black text-base text-slate-950 bg-white/85 border border-slate-500 focus:border-blue-700 focus:outline-hidden"
                      />
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Live Formula Display Box */}
            <div className="bg-white border-2 border-red-600 px-2.5 py-2 text-center font-serif font-black text-xs sm:text-sm text-slate-950 tracking-tight">
              {liveFormulaText}
            </div>

            {/* Informasi Box (Cyan background, red border) */}
            <div className="bg-[#00ffff] border-2 border-red-600 p-3 text-[11px] leading-relaxed text-slate-950">
              <div className="font-bold text-blue-900 text-xs mb-1">Informasi</div>
              <p>
                Penilaian pada aplikasi ini dikelompokkan menjadi 3 (tiga) yaitu Penilaian untuk setiap Tujuan Pembelajaran, Penilaian setiap Lingkup Materi, dan Penilaian Akhir Semester. Nama Penilaian pada Tabel <strong>&quot;BOBOT NILAI RAPOR&quot;</strong> dapat <strong>diubah/diganti</strong> sesuai dengan kebijakan yang berlaku di satuan pendidikan masing-masing. Pengguna dapat mengganti Nama penilaian pada setiap Tujuan Pembelajaran (baris pertama pada tabel <strong>&quot;BOBOT NILAI RAPOR&quot;</strong>) menjadi <strong>FORMATIF</strong> atau <strong>SUMATIF TUJUAN PEMBELAJARAN</strong>. Yang perlu menjadi perhatian adalah jika nama penilaian pada Tujuan Pembelajaran diubah menjadi <strong>FORMATIF</strong> maka Bobot Nilai Rapor pada baris tersebut otomatis diubah menjadi nol (0), sebab nilai FORMATIF sesuai petunjuk teknis penilaian tidak diolah menjadi nilai RAPOR, hanya untuk pertimbangan <strong>DESKRIPSI RAPOR</strong> saja. Jika Bapak/Ibu menginginkan nilai pada Setiap Tujuan Pembelajaran tetap diolah menjadi nilai RAPOR maka nama penilaian pada Tujuan Pembelajaran ditetapkan <strong>SUMATIF TUJUAN PEMBELAJARAN</strong>.
              </p>
            </div>
          </div>

          {/* COLUMN 3: KKTP TABLE + LIVE FORMULA SIMULATION */}
          <div className="space-y-2.5">
            <div className="bg-white border-2 border-red-600 overflow-x-auto">
              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr className="bg-black text-white font-serif">
                    <th className="border border-red-600 px-2 py-2 text-center w-28 font-black">
                      <div className="inline-flex items-center justify-center gap-1 leading-tight">
                        <span>RENTANG NILAI</span>
                        <button
                          type="button"
                          onClick={() => setActiveHelp(activeHelp === 'kktp' ? 'none' : 'kktp')}
                          title="Bantuan KKTP"
                          className="text-purple-400 hover:text-purple-300 cursor-pointer shrink-0"
                        >
                          <HelpCircle size={13} />
                        </button>
                      </div>
                    </th>
                    <th className="border border-red-600 px-2 py-2 text-center w-36 font-black">
                      PREDIKAT
                    </th>
                    <th className="border border-red-600 px-2 py-2 text-center font-black">
                      KALIMAT AWAL DESKRIPSI RAPOR
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {settings.kktpRanges.map((row, idx) => (
                    <tr key={row.id} className="bg-white hover:bg-amber-50/40">
                      <td className="border border-red-600 px-1.5 py-2 text-center">
                        <div className="inline-flex items-center justify-center gap-1 font-serif font-black text-slate-950">
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={row.min}
                            onChange={(e) => handleUpdateKktpRow(idx, 'min', e.target.value)}
                            aria-label={`Nilai minimum ${row.labelPredikat}`}
                            className="w-11 px-1 py-0.5 text-center font-mono tabular-nums font-black border border-slate-300 focus:border-blue-600 focus:outline-hidden bg-white"
                          />
                          <span>-</span>
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={row.max}
                            onChange={(e) => handleUpdateKktpRow(idx, 'max', e.target.value)}
                            aria-label={`Nilai maksimum ${row.labelPredikat}`}
                            className="w-11 px-1 py-0.5 text-center font-mono tabular-nums font-black border border-slate-300 focus:border-blue-600 focus:outline-hidden bg-white"
                          />
                        </div>
                      </td>
                      <td className="border border-red-600 px-1.5 py-2">
                        <input
                          type="text"
                          value={row.labelPredikat}
                          onChange={(e) => handleUpdateKktpRow(idx, 'labelPredikat', e.target.value)}
                          aria-label={`Predikat ${row.id}`}
                          className="w-full px-1.5 py-0.5 font-serif font-bold italic text-sm text-slate-950 bg-transparent border border-transparent focus:border-blue-500 focus:bg-white focus:outline-hidden"
                        />
                      </td>
                      <td className="border border-red-600 px-1.5 py-2">
                        <input
                          type="text"
                          value={row.kalimatAwal}
                          onChange={(e) => handleUpdateKktpRow(idx, 'kalimatAwal', e.target.value)}
                          aria-label={`Kalimat awal ${row.id}`}
                          className="w-full px-1.5 py-0.5 font-serif italic text-xs text-slate-950 bg-transparent border border-transparent focus:border-blue-500 focus:bg-white focus:outline-hidden"
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Live Rumus & Kalimat Deskripsi Simulation Card */}
            <div className="bg-white border-2 border-red-600 p-3 space-y-2.5 text-xs">
              <div className="flex flex-wrap items-center justify-between gap-2 pb-2 border-b border-slate-200">
                <div className="flex items-center gap-1.5 font-black text-slate-900">
                  <Calculator size={15} className="text-blue-700" />
                  <span>Pratinjau Rumus Bobot &amp; Deskripsi Otomatis</span>
                </div>
                <StudentFilterSelect compact />
              </div>

              <div className="grid grid-cols-4 gap-1.5 text-center">
                <div className="bg-purple-50 border border-purple-200 p-1.5 rounded-xs">
                  <div className="text-[10px] font-bold text-purple-900">
                    {settings.bobot.namaPenilaianTp === 'FORMATIF' ? 'FORMATIF' : 'S. TP'}
                  </div>
                  <div className="font-mono font-black text-sm text-slate-950">
                    {previewResult.sampleTp} <span className="text-[10px] font-normal">(×{activeBobotTp})</span>
                  </div>
                </div>
                <div className="bg-amber-50 border border-amber-300 p-1.5 rounded-xs">
                  <div className="text-[10px] font-bold text-amber-900">S. LM</div>
                  <div className="font-mono font-black text-sm text-slate-950">
                    {previewResult.sampleLm} <span className="text-[10px] font-normal">(×{activeBobotLm})</span>
                  </div>
                </div>
                <div className="bg-cyan-50 border border-cyan-300 p-1.5 rounded-xs">
                  <div className="text-[10px] font-bold text-cyan-900">S. AS</div>
                  <div className="font-mono font-black text-sm text-slate-950">
                    {previewResult.sampleSas} <span className="text-[10px] font-normal">(×{activeBobotSas})</span>
                  </div>
                </div>
                <div className="bg-emerald-50 border border-emerald-400 p-1.5 rounded-xs">
                  <div className="text-[10px] font-bold text-emerald-900">NA RAPOR</div>
                  <div className="font-mono font-black text-sm text-emerald-900">
                    {previewResult.weighted} ({previewResult.matchedKktp.labelPredikat})
                  </div>
                </div>
              </div>

              <div className="bg-slate-50 border border-slate-300 p-2 rounded-xs">
                <div className="text-[10px] font-bold text-slate-500 uppercase mb-0.5">
                  Contoh Hasil Deskripsi Rapor ({previewResult.callName || 'Siswa'}):
                </div>
                <div className="font-serif italic text-slate-900 leading-snug">
                  &ldquo;{previewResult.sampleSentence}&rdquo;
                </div>
              </div>

              <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
                <span className="text-[11px] text-slate-600">
                  Klik tombol simpan agar seluruh <strong>Input Nilai</strong>, <strong>TP</strong> &amp; <strong>Cetak Rapor</strong> memakai rumus &amp; KKM terbaru.
                </span>
                <button
                  type="button"
                  onClick={handleConnectFormulaToAllSubjects}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white font-bold text-[11px] rounded-xs cursor-pointer"
                >
                  <Save size={13} />
                  <span>Simpan &amp; Hubungkan Semua</span>
                </button>
              </div>
            </div>

            {/* Optional SK Metadata Editor (toggleable via printer button) */}
            {showSkModal && (
              <div className="bg-white border-2 border-slate-950 p-3 space-y-2 text-xs">
                <div className="flex items-center justify-between">
                  <span className="font-black text-slate-900">Pengaturan SK Penetapan KKM (Untuk Cetak)</span>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => window.print()}
                      className="inline-flex items-center gap-1 px-2.5 py-1 bg-slate-900 text-white font-bold rounded-xs cursor-pointer"
                    >
                      <Printer size={12} /> Cetak SK Resmi
                    </button>
                    <button
                      type="button"
                      onClick={() => setShowSkModal(false)}
                      className="p-1 text-slate-500 hover:text-slate-900 cursor-pointer"
                    >
                      <X size={14} />
                    </button>
                  </div>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Nomor SK</label>
                    <input
                      type="text"
                      value={skNumber}
                      onChange={(e) => setSkNumber(e.target.value)}
                      className="w-full px-2 py-1 border border-slate-300 font-mono text-[11px]"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-0.5">Tanggal Penetapan</label>
                    <input
                      type="text"
                      value={skDate}
                      onChange={(e) => setSkDate(e.target.value)}
                      className="w-full px-2 py-1 border border-slate-300 text-[11px]"
                    />
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* PRINTABLE OFFICIAL SHEET FOR PRINTING */}
      <div className="hidden print:block print:w-full print:p-6 bg-white text-slate-900">
        <div className="text-center pb-4 border-b-2 border-slate-900 mb-6">
          <div className="text-sm font-bold uppercase">
            {db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'}
          </div>
          <div className="text-sm font-bold uppercase">
            {db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'}
          </div>
          <h1 className="text-xl font-black uppercase mt-1">
            {db.sekolah.namaSekolah || 'SD NEGERI BALIL'}
          </h1>
          <p className="text-xs text-slate-600">
            Alamat: {db.sekolah.alamat}, Kab. {db.sekolah.kabupaten}, Provinsi {db.sekolah.provinsi}
          </p>
        </div>

        <div className="text-center mb-5">
          <h2 className="text-base font-black uppercase underline">
            SURAT KEPUTUSAN PENETAPAN KKM, BOBOT NILAI RAPOR &amp; KKTP
          </h2>
          <p className="text-xs font-bold mt-1">Nomor: {skNumber}</p>
          <p className="text-xs">
            Tahun Pelajaran: {db.sekolah.tahunPelajaran} • Rumus: {liveFormulaText}
          </p>
        </div>

        <table className="w-full text-xs border-collapse border border-slate-900 mb-5">
          <thead>
            <tr className="bg-slate-100 text-center font-bold">
              <th className="border border-slate-900 p-2 w-10">No</th>
              <th className="border border-slate-900 p-2 text-left">Muatan / Mata Pelajaran</th>
              <th className="border border-slate-900 p-2 w-28">Kelompok</th>
              <th className="border border-slate-900 p-2 w-20">Nilai KKM</th>
            </tr>
          </thead>
          <tbody>
            {previewMapels.map((m, i) => (
              <tr key={m.kode || m.mataPelajaran || i}>
                <td className="border border-slate-900 p-1.5 text-center font-mono">{i + 1}</td>
                <td className="border border-slate-900 p-1.5 font-bold">{m.mataPelajaran}</td>
                <td className="border border-slate-900 p-1.5 text-center">{m.kelompok}</td>
                <td className="border border-slate-900 p-1.5 text-center font-bold font-mono">
                  {m.kkm ?? 75}
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        <table className="w-full text-xs border-collapse border border-slate-900 mb-6">
          <thead>
            <tr className="bg-slate-100 text-center font-bold">
              <th className="border border-slate-900 p-1.5 w-28">Rentang Nilai</th>
              <th className="border border-slate-900 p-1.5 w-36">Predikat</th>
              <th className="border border-slate-900 p-1.5 text-left">Kalimat Awal Deskripsi Rapor</th>
            </tr>
          </thead>
          <tbody>
            {settings.kktpRanges.map((r) => (
              <tr key={r.id}>
                <td className="border border-slate-900 p-1.5 text-center font-mono font-bold">
                  {r.min} - {r.max}
                </td>
                <td className="border border-slate-900 p-1.5 font-bold italic">{r.labelPredikat}</td>
                <td className="border border-slate-900 p-1.5 italic">{r.kalimatAwal}</td>
              </tr>
            ))}
          </tbody>
        </table>

        <div className="grid grid-cols-2 text-xs pt-4">
          <div className="text-center">
            <p>Mengetahui,</p>
            <p className="font-bold">Guru Kelas / Tim Kurikulum</p>
            <div className="h-20"></div>
            <p className="font-bold underline">{db.sekolah.waliKelas || 'Selpius Nabyal, S.Pd.Gr'}</p>
            <p className="font-mono">NIP: {db.sekolah.nipWaliKelas || '-'}</p>
          </div>
          <div className="text-center">
            <p>
              {db.sekolah.tempatRapor || 'Oksibil'}, {skDate}
            </p>
            <p className="font-bold">Kepala Sekolah,</p>
            <div className="h-20"></div>
            <p className="font-bold underline">
              {db.sekolah.namaKepalaSekolah || 'Selpius Nabyal, S.Pd.Gr.'}
            </p>
            <p className="font-mono">NIP: {db.sekolah.nipKepalaSekolah || '-'}</p>
          </div>
        </div>
      </div>
    </div>
  );
};
