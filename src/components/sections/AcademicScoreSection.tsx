import React, { useState, useEffect, useMemo } from 'react';
import {
  FileSpreadsheet,
  Save,
  CheckCircle2,
  RefreshCw,
  BookOpen,
  FolderDown,
  Eye,
  Zap,
  Download,
  Columns3,
  Edit3,
  X
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { AcademicScore } from '../../types';
import { StandardActionToolbar } from '../common/StandardActionToolbar';
import { SaveStatusIndicator } from '../common/SaveStatusIndicator';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { useSaveFeedback } from '../../hooks/useSaveFeedback';
import {
  generateReportDescriptionFromTP,
  getMatchedTpsForSubjectAndStudent,
  calculateKurikulumMerdekaScores,
  hasKurikulumMerdekaInputs,
  resolveTpSlots,
  TP_COLUMN_KEYS,
  LM_COLUMN_KEYS,
  getPredikatFromScore,
  getCapaianFromScore,
  clampScore,
  formatStudentCallName
} from '../../utils/tpDescriptionHelper';
import {
  exportSubjectScoresToExcel,
  exportSubjectScoresToCsv,
  exportClassComprehensiveWorkbook
} from '../../utils/exportHelper';
import * as XLSX from 'xlsx';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import { matchesStudentScore, isPlaceholderNisn } from '../../utils/studentSyncHelper';

const RELIGION_TABS = [
  { key: 'Hindu', label: 'HINDU', bg: 'bg-fuchsia-500 hover:bg-fuchsia-600 text-white border-fuchsia-300' },
  { key: 'Islam', label: 'ISLAM', bg: 'bg-purple-600 hover:bg-purple-700 text-white border-purple-300' },
  { key: 'Kristen', label: 'KRISTEN', bg: 'bg-blue-600 hover:bg-blue-700 text-white border-blue-300' },
  { key: 'Katolik', label: 'KATOLIK', bg: 'bg-slate-200 hover:bg-slate-300 text-slate-900 border-slate-400' },
  { key: 'Buddha', label: 'BUDHA', bg: 'bg-emerald-500 hover:bg-emerald-600 text-white border-emerald-300' },
  { key: 'Khonghucu', label: 'KONGHUCU', bg: 'bg-yellow-400 hover:bg-yellow-500 text-slate-950 border-yellow-200' }
] as const;

const TP_HEADER_COLORS = [
  'bg-pink-300 text-slate-950',
  'bg-yellow-300 text-slate-950',
  'bg-sky-400 text-slate-950',
  'bg-cyan-300 text-slate-950',
  'bg-blue-500 text-white',
  'bg-fuchsia-300 text-slate-950',
  'bg-amber-300 text-slate-950',
  'bg-teal-300 text-slate-950',
  'bg-lime-300 text-slate-950',
  'bg-rose-300 text-slate-950',
  'bg-indigo-300 text-slate-950',
  'bg-orange-300 text-slate-950'
];

const LM_HEADER_COLORS = [
  'bg-amber-400 text-slate-950',
  'bg-cyan-200 text-slate-950',
  'bg-yellow-300 text-slate-950',
  'bg-sky-400 text-slate-950',
  'bg-pink-300 text-slate-950',
  'bg-lime-300 text-slate-950',
  'bg-indigo-200 text-slate-950',
  'bg-teal-300 text-slate-950'
];

function matchesReligion(studentAgama: string | undefined, filterAgama: string): boolean {
  if (!filterAgama) return true;
  const s = (studentAgama || '').trim().toLowerCase();
  const f = filterAgama.trim().toLowerCase();
  if (f === 'khonghucu' || f === 'konghucu') {
    return s.includes('khonghucu') || s.includes('konghucu');
  }
  if (f === 'buddha' || f === 'budha') {
    return s.includes('buddha') || s.includes('budha');
  }
  return s.includes(f);
}

export const AcademicScoreSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    selectedKelas,
    selectedSemester,
    filteredStudent,
    saveNilaiAkademik,
    bulkSaveNilai,
    saveLingkupMateriTitles,
    syncAllScoresWithTPDescriptions,
    showSaveNotification,
    setActiveMenu,
    openReportPreview
  } = appContext;
  const { status: saveStatus, triggerSave, lastSavedTime } = useSaveFeedback();

  const [selectedMapel, setSelectedMapel] = useState<string>(
    db.mataPelajaran[0]?.mataPelajaran || 'Pendidikan Agama dan Budi Pekerti'
  );
  const [agamaFilter, setAgamaFilter] = useState<string>('');
  const [showHiddenColumns, setShowHiddenColumns] = useState<boolean>(false);
  const [editingLmKey, setEditingLmKey] = useState<string | null>(null);
  const [editingLmDraft, setEditingLmDraft] = useState<string>('');
  const [activeTpInfoIdx, setActiveTpInfoIdx] = useState<number | null>(null);

  const siswaKelas = useMemo(
    () => db.siswa.filter((s) => s.kelas === selectedKelas),
    [db.siswa, selectedKelas]
  );

  const displayedSiswa = useMemo(() => {
    let base = filteredStudent
      ? siswaKelas.filter((s) => s.idSiswa === filteredStudent.idSiswa || matchesStudentScore(s, filteredStudent))
      : siswaKelas;
    if (agamaFilter) {
      base = base.filter((s) => matchesReligion(s.agama, agamaFilter));
    }
    return base;
  }, [siswaKelas, filteredStudent, agamaFilter]);

  const activeTpList = useMemo(() => {
    const refAgama = agamaFilter || displayedSiswa[0]?.agama || '';
    return getMatchedTpsForSubjectAndStudent(
      selectedMapel,
      selectedKelas,
      db.tujuanPembelajaran,
      refAgama
    );
  }, [selectedMapel, selectedKelas, db.tujuanPembelajaran, agamaFilter, displayedSiswa]);

  const activeTpSlots = useMemo(() => resolveTpSlots(activeTpList), [activeTpList]);

  const lmScopeKey = `${selectedKelas}::${selectedMapel}`;
  const lmTitles = db.lingkupMateriTitles?.[lmScopeKey] || {};

  const visibleTpKeys = useMemo(() => {
    if (showHiddenColumns) {
      return [...TP_COLUMN_KEYS];
    }
    const highestTpSlot = activeTpSlots.reduce((max, tp, idx) => (tp ? idx + 1 : max), 8);
    const baseCount = Math.min(12, Math.max(8, highestTpSlot, activeTpList.length));
    return TP_COLUMN_KEYS.slice(0, baseCount);
  }, [showHiddenColumns, activeTpList.length, activeTpSlots]);

  const visibleLmKeys = useMemo(() => {
    if (showHiddenColumns) {
      return [...LM_COLUMN_KEYS];
    }
    const highestLmSlot = LM_COLUMN_KEYS.reduce(
      (max, k, idx) => (lmTitles[k] ? idx + 1 : max),
      4
    );
    return LM_COLUMN_KEYS.slice(0, Math.min(8, Math.max(4, highestLmSlot)));
  }, [showHiddenColumns, lmTitles]);

  const buildInitialRow = (s: (typeof siswaKelas)[number], mapelName: string): AcademicScore => {
    const existing = db.nilaiAkademik.find(
      (n) => matchesStudentScore(n, s) && n.mataPelajaran === mapelName && n.semester === selectedSemester
    );
    const mapelObj = db.mataPelajaran.find((m) => m.mataPelajaran === mapelName);
    const mapelSlug = mapelName.replace(/[^a-zA-Z0-9]/g, '').slice(0, 6);
    const matchedTps = getMatchedTpsForSubjectAndStudent(
      mapelName,
      selectedKelas,
      db.tujuanPembelajaran,
      s.agama
    );

    if (existing) {
      if (hasKurikulumMerdekaInputs(existing)) {
        const subjectLmTitles = db.lingkupMateriTitles?.[`${selectedKelas}::${mapelName}`];
        const callName = formatStudentCallName(s.namaSiswa, s.bukuInduk?.namaPanggilan);
        const km = calculateKurikulumMerdekaScores(
          existing,
          matchedTps,
          mapelName,
          subjectLmTitles,
          db.kkmSettings,
          callName
        );
        const effectiveNa = existing.nilaiAkhir > 0 ? existing.nilaiAkhir : km.nilaiAkhir;
        return {
          ...existing,
          idSiswa: s.idSiswa,
          nisn: s.nisn,
          namaSiswa: s.namaSiswa,
          kelas: selectedKelas,
          kkm: mapelObj?.kkm ?? existing.kkm ?? 75,
          naSumatifTp: km.naSumatifTp,
          naSumatifLm: km.naSumatifLm,
          naSumatifSas: km.naSumatifSas,
          jumlahNilai: km.jumlahNilai,
          nilaiAkhir: effectiveNa,
          predikat: getPredikatFromScore(effectiveNa, db.kkmSettings?.kktpRanges),
          capaianKompetensi: getCapaianFromScore(effectiveNa, db.kkmSettings?.kktpRanges),
          deskripsiTercapai: km.deskripsiTercapai || existing.deskripsiTercapai,
          deskripsiPerluPeningkatan: km.deskripsiPerluPeningkatan || existing.deskripsiPerluPeningkatan,
          deskripsiRapor: km.deskripsiRapor || existing.deskripsiRapor
        };
      }
      return {
        ...existing,
        idSiswa: s.idSiswa,
        nisn: s.nisn,
        namaSiswa: s.namaSiswa,
        kelas: selectedKelas,
        kkm: mapelObj?.kkm ?? existing.kkm ?? 75
      };
    }

    return {
      idNilai: `N_${s.idSiswa}_${mapelSlug}_${selectedSemester}`,
      idSiswa: s.idSiswa,
      tahunPelajaran: db.sekolah.tahunPelajaran,
      semester: selectedSemester,
      kelas: selectedKelas,
      nisn: s.nisn,
      namaSiswa: s.namaSiswa,
      mataPelajaran: mapelName,
      kkm: mapelObj?.kkm ?? 75,
      tpScores: {},
      lmScores: {},
      sasNonTes: undefined,
      sasTes: undefined,
      naSumatifTp: undefined,
      naSumatifLm: undefined,
      naSumatifSas: undefined,
      jumlahNilai: 0,
      nilaiAkhir: 0,
      predikat: 'Perlu Bimbingan',
      capaianKompetensi: 'Perlu Pendampingan',
      deskripsiTercapai: '',
      deskripsiPerluPeningkatan: '',
      deskripsiRapor: ''
    };
  };

  const [rows, setRows] = useState<Record<string, AcademicScore>>(() => {
    const initialMap: Record<string, AcademicScore> = {};
    siswaKelas.forEach((s) => {
      initialMap[s.idSiswa] = buildInitialRow(s, selectedMapel);
    });
    return initialMap;
  });

  useEffect(() => {
    setRows(() => {
      const next: Record<string, AcademicScore> = {};
      siswaKelas.forEach((s) => {
        next[s.idSiswa] = buildInitialRow(s, selectedMapel);
      });
      return next;
    });
  }, [
    db.siswa,
    db.nilaiAkademik,
    db.tujuanPembelajaran,
    db.lingkupMateriTitles,
    db.kkmSettings,
    db.mataPelajaran,
    selectedMapel,
    selectedSemester,
    selectedKelas
  ]);

  const handleMapelChange = (newMapel: string) => {
    setSelectedMapel(newMapel);
    if (!newMapel.toLowerCase().includes('agama')) {
      setAgamaFilter('');
    }
    const initialMap: Record<string, AcademicScore> = {};
    siswaKelas.forEach((s) => {
      initialMap[s.idSiswa] = buildInitialRow(s, newMapel);
    });
    setRows(initialMap);
  };

  const handleReligionTabClick = (religionKey: string) => {
    const nextFilter = agamaFilter === religionKey ? '' : religionKey;
    setAgamaFilter(nextFilter);
    if (nextFilter) {
      const agamaSubject = db.mataPelajaran.find((m) =>
        m.mataPelajaran.toLowerCase().includes('agama')
      );
      if (agamaSubject && selectedMapel !== agamaSubject.mataPelajaran) {
        setSelectedMapel(agamaSubject.mataPelajaran);
      }
    }
  };

  const applyKmUpdate = (studentObj: (typeof siswaKelas)[number], draft: AcademicScore) => {
    const matchedTps = getMatchedTpsForSubjectAndStudent(
      selectedMapel,
      selectedKelas,
      db.tujuanPembelajaran,
      studentObj.agama
    );
    const callName = formatStudentCallName(studentObj.namaSiswa, studentObj.bukuInduk?.namaPanggilan);
    const km = calculateKurikulumMerdekaScores(
      draft,
      matchedTps,
      selectedMapel,
      lmTitles,
      db.kkmSettings,
      callName
    );
    const fallbackDesc =
      km.nilaiAkhir > 0 && !km.deskripsiTercapai
        ? generateReportDescriptionFromTP(
            selectedMapel,
            selectedKelas,
            km.nilaiAkhir,
            db.tujuanPembelajaran,
            studentObj.namaSiswa,
            studentObj.agama,
            db.kkmSettings,
            studentObj.bukuInduk?.namaPanggilan
          )
        : null;

    const tertinggi = km.deskripsiTercapai || fallbackDesc?.deskripsiTercapai || '';
    const terendah = km.deskripsiPerluPeningkatan || fallbackDesc?.deskripsiPerluPeningkatan || '';
    const combined =
      km.deskripsiRapor ||
      (tertinggi && terendah ? `${tertinggi} ${terendah}` : tertinggi || fallbackDesc?.deskripsiRapor || '');

    return {
      ...draft,
      naSumatifTp: km.naSumatifTp,
      naSumatifLm: km.naSumatifLm,
      naSumatifSas: km.naSumatifSas,
      jumlahNilai: km.jumlahNilai,
      nilaiAkhir: km.nilaiAkhir,
      predikat: getPredikatFromScore(km.nilaiAkhir, db.kkmSettings?.kktpRanges),
      capaianKompetensi: getCapaianFromScore(km.nilaiAkhir, db.kkmSettings?.kktpRanges),
      deskripsiTercapai: km.hasKmInputs ? tertinggi : '',
      deskripsiPerluPeningkatan: km.hasKmInputs ? terendah : '',
      deskripsiRapor: km.hasKmInputs ? combined : ''
    };
  };

  const updateTpCell = (studentId: string, tpKey: string, value: string) => {
    const studentObj = siswaKelas.find((s) => s.idSiswa === studentId);
    if (!studentObj) return;
    const current = rows[studentId] || buildInitialRow(studentObj, selectedMapel);
    const rawStr = value.trim();
    const parsed = rawStr === '' ? undefined : clampScore(rawStr);

    const nextTpScores = { ...(current.tpScores || {}), [tpKey]: parsed };
    if (parsed === undefined) {
      delete nextTpScores[tpKey];
    }

    const updated = applyKmUpdate(studentObj, {
      ...current,
      idSiswa: studentObj.idSiswa,
      nisn: studentObj.nisn,
      namaSiswa: studentObj.namaSiswa,
      tpScores: nextTpScores
    });

    setRows((prev) => ({ ...prev, [studentId]: updated }));
    saveNilaiAkademik(updated, true);
  };

  const updateLmCell = (studentId: string, lmKey: string, value: string) => {
    const studentObj = siswaKelas.find((s) => s.idSiswa === studentId);
    if (!studentObj) return;
    const current = rows[studentId] || buildInitialRow(studentObj, selectedMapel);
    const rawStr = value.trim();
    const parsed = rawStr === '' ? undefined : clampScore(rawStr);

    const nextLmScores = { ...(current.lmScores || {}), [lmKey]: parsed };
    if (parsed === undefined) {
      delete nextLmScores[lmKey];
    }

    const updated = applyKmUpdate(studentObj, {
      ...current,
      idSiswa: studentObj.idSiswa,
      nisn: studentObj.nisn,
      namaSiswa: studentObj.namaSiswa,
      lmScores: nextLmScores
    });

    setRows((prev) => ({ ...prev, [studentId]: updated }));
    saveNilaiAkademik(updated, true);
  };

  const updateSasCell = (studentId: string, field: 'sasNonTes' | 'sasTes', value: string) => {
    const studentObj = siswaKelas.find((s) => s.idSiswa === studentId);
    if (!studentObj) return;
    const current = rows[studentId] || buildInitialRow(studentObj, selectedMapel);
    const rawStr = value.trim();
    const parsed = rawStr === '' ? undefined : clampScore(rawStr);

    const updated = applyKmUpdate(studentObj, {
      ...current,
      idSiswa: studentObj.idSiswa,
      nisn: studentObj.nisn,
      namaSiswa: studentObj.namaSiswa,
      [field]: parsed
    });

    setRows((prev) => ({ ...prev, [studentId]: updated }));
    saveNilaiAkademik(updated, true);
  };

  const updateManualField = (
    studentId: string,
    field: 'nilaiAkhir' | 'deskripsiTercapai' | 'deskripsiPerluPeningkatan' | 'deskripsiRapor',
    value: string
  ) => {
    const studentObj = siswaKelas.find((s) => s.idSiswa === studentId);
    if (!studentObj) return;
    const current = rows[studentId] || buildInitialRow(studentObj, selectedMapel);
    let updated: AcademicScore = {
      ...current,
      idSiswa: studentObj.idSiswa,
      nisn: studentObj.nisn,
      namaSiswa: studentObj.namaSiswa
    };

    if (field === 'nilaiAkhir') {
      const rawStr = value.trim();
      const na = rawStr === '' ? 0 : clampScore(rawStr);
      const tpRes =
        na > 0
          ? generateReportDescriptionFromTP(
              selectedMapel,
              selectedKelas,
              na,
              db.tujuanPembelajaran,
              studentObj.namaSiswa,
              studentObj.agama,
              db.kkmSettings,
              studentObj.bukuInduk?.namaPanggilan
            )
          : { deskripsiTercapai: '', deskripsiPerluPeningkatan: '', deskripsiRapor: '' };
      updated = {
        ...updated,
        nilaiAkhir: na,
        predikat: getPredikatFromScore(na, db.kkmSettings?.kktpRanges),
        capaianKompetensi: getCapaianFromScore(na, db.kkmSettings?.kktpRanges),
        deskripsiTercapai: updated.deskripsiTercapai || tpRes.deskripsiTercapai,
        deskripsiPerluPeningkatan: updated.deskripsiPerluPeningkatan || tpRes.deskripsiPerluPeningkatan,
        deskripsiRapor: updated.deskripsiRapor || tpRes.deskripsiRapor
      };
    } else if (field === 'deskripsiRapor') {
      updated = {
        ...updated,
        deskripsiRapor: value
      };
    } else {
      updated = {
        ...updated,
        [field]: value
      };
      const tertinggi = field === 'deskripsiTercapai' ? value : updated.deskripsiTercapai;
      const terendah = field === 'deskripsiPerluPeningkatan' ? value : updated.deskripsiPerluPeningkatan;
      updated.deskripsiRapor = [tertinggi, terendah].filter(Boolean).join(' ').trim();
    }

    setRows((prev) => ({ ...prev, [studentId]: updated }));
    saveNilaiAkademik(updated, true);
  };

  const handleSaveAll = () => {
    triggerSave(() => {
      const listToSave = Object.values(rows);
      bulkSaveNilai(listToSave);
    });
  };

  const handleRecalculateAll = () => {
    const recalculated: Record<string, AcademicScore> = {};
    Object.keys(rows).forEach((studentId) => {
      const item = rows[studentId];
      const studentObj = siswaKelas.find((s) => s.idSiswa === studentId);
      if (!studentObj) return;
      if (hasKurikulumMerdekaInputs(item)) {
        recalculated[studentId] = applyKmUpdate(studentObj, item);
      } else {
        const na = item.nilaiAkhir || 0;
        const tpRes =
          na > 0
            ? generateReportDescriptionFromTP(
                selectedMapel,
                selectedKelas,
                na,
                db.tujuanPembelajaran,
                studentObj.namaSiswa,
                studentObj.agama,
                db.kkmSettings,
                studentObj.bukuInduk?.namaPanggilan
              )
            : { deskripsiTercapai: '', deskripsiPerluPeningkatan: '', deskripsiRapor: '' };
        recalculated[studentId] = {
          ...item,
          idSiswa: studentObj.idSiswa,
          nisn: studentObj.nisn,
          namaSiswa: studentObj.namaSiswa,
          predikat: getPredikatFromScore(na, db.kkmSettings?.kktpRanges),
          capaianKompetensi: getCapaianFromScore(na, db.kkmSettings?.kktpRanges),
          deskripsiTercapai: tpRes.deskripsiTercapai,
          deskripsiPerluPeningkatan: tpRes.deskripsiPerluPeningkatan,
          deskripsiRapor: tpRes.deskripsiRapor
        };
      }
    });
    setRows(recalculated);
    triggerSave(() => {
      bulkSaveNilai(Object.values(recalculated));
    });
  };

  const handleConnectAllStudentsToAllSubjects = () => {
    const allSubjects = db.mataPelajaran;
    const targetStudents = siswaKelas.length > 0 ? siswaKelas : db.siswa;
    const scoresToSave: AcademicScore[] = [];

    targetStudents.forEach((student) => {
      allSubjects.forEach((mapel) => {
        const existing = db.nilaiAkademik.find(
          (n) =>
            matchesStudentScore(n, student) &&
            n.mataPelajaran === mapel.mataPelajaran &&
            n.kelas === student.kelas &&
            n.semester === selectedSemester
        );
        const matchedTps = getMatchedTpsForSubjectAndStudent(
          mapel.mataPelajaran,
          student.kelas,
          db.tujuanPembelajaran,
          student.agama
        );
        const useKm = existing ? hasKurikulumMerdekaInputs(existing) : false;
        const subjectLmTitles = db.lingkupMateriTitles?.[`${student.kelas}::${mapel.mataPelajaran}`];
        const callName = formatStudentCallName(student.namaSiswa, student.bukuInduk?.namaPanggilan);
        const km = useKm && existing
          ? calculateKurikulumMerdekaScores(
              existing,
              matchedTps,
              mapel.mataPelajaran,
              subjectLmTitles,
              db.kkmSettings,
              callName
            )
          : null;
        const effectiveNA = km ? km.nilaiAkhir : (existing?.nilaiAkhir ?? 0);
        const tpRes =
          effectiveNA > 0 && !km?.deskripsiTercapai
            ? generateReportDescriptionFromTP(
                mapel.mataPelajaran,
                student.kelas,
                effectiveNA,
                db.tujuanPembelajaran,
                student.namaSiswa,
                student.agama,
                db.kkmSettings,
                student.bukuInduk?.namaPanggilan
              )
            : null;

        scoresToSave.push({
          idNilai:
            existing?.idNilai ||
            `N_${student.idSiswa}_${mapel.mataPelajaran.replace(/[^a-zA-Z0-9]/g, '').slice(0, 6)}_${selectedSemester}`,
          idSiswa: student.idSiswa,
          tahunPelajaran: db.sekolah.tahunPelajaran,
          semester: selectedSemester,
          kelas: student.kelas,
          nisn: student.nisn,
          namaSiswa: student.namaSiswa,
          mataPelajaran: mapel.mataPelajaran,
          kkm: mapel.kkm ?? existing?.kkm ?? 75,
          tpScores: existing?.tpScores,
          lmScores: existing?.lmScores,
          sasNonTes: existing?.sasNonTes,
          sasTes: existing?.sasTes,
          naSumatifTp: km?.naSumatifTp,
          naSumatifLm: km?.naSumatifLm,
          naSumatifSas: km?.naSumatifSas,
          jumlahNilai: km ? km.jumlahNilai : (existing?.jumlahNilai ?? 0),
          nilaiAkhir: effectiveNA,
          predikat: getPredikatFromScore(effectiveNA, db.kkmSettings?.kktpRanges),
          capaianKompetensi: getCapaianFromScore(effectiveNA, db.kkmSettings?.kktpRanges),
          deskripsiTercapai: km?.deskripsiTercapai || tpRes?.deskripsiTercapai || existing?.deskripsiTercapai || '',
          deskripsiPerluPeningkatan: km?.deskripsiPerluPeningkatan || tpRes?.deskripsiPerluPeningkatan || existing?.deskripsiPerluPeningkatan || '',
          deskripsiRapor: km?.deskripsiRapor || tpRes?.deskripsiRapor || existing?.deskripsiRapor || ''
        });
      });
    });

    bulkSaveNilai(scoresToSave);
    showSaveNotification(
      `Berhasil menghubungkan ${targetStudents.length} siswa ke seluruh ${allSubjects.length} mata pelajaran!`
    );
    handleMapelChange(selectedMapel);
  };

  const handleExportExcel = () => {
    exportSubjectScoresToExcel(Object.values(rows), selectedMapel, selectedKelas, selectedSemester, db.sekolah);
  };

  const handleExportCsv = () => {
    exportSubjectScoresToCsv(Object.values(rows), selectedMapel, selectedKelas, selectedSemester, db.sekolah);
  };

  const handleExportFullWorkbook = () => {
    exportClassComprehensiveWorkbook(db, selectedKelas, selectedSemester);
  };

  const handleImportFile = async (file: File) => {
    try {
      const result = await parseAndImportTemplate(file, 'template_nilai_akademik', appContext);
      if (result.success) {
        showSaveNotification(result.message);
      } else {
        showSaveNotification(`Gagal impor: ${result.message}`);
      }
    } catch (e: any) {
      console.error('Import error', e);
      showSaveNotification('Gagal mengimpor file nilai: ' + (e?.message || 'Format berkas tidak sesuai'));
    }
  };

  const handleClearScores = () => {
    const clearedList: AcademicScore[] = [];
    setRows((prev) => {
      const resetMap: Record<string, AcademicScore> = {};
      Object.entries(prev).forEach(([studentId, score]) => {
        const cleared: AcademicScore = {
          ...score,
          formatif: undefined,
          sumatif: undefined,
          ph1: undefined,
          ph2: undefined,
          ph3: undefined,
          pts: undefined,
          pas: undefined,
          praktik: undefined,
          tpScores: {},
          lmScores: {},
          sasNonTes: undefined,
          sasTes: undefined,
          naSumatifTp: undefined,
          naSumatifLm: undefined,
          naSumatifSas: undefined,
          jumlahNilai: 0,
          nilaiAkhir: 0,
          predikat: 'Perlu Bimbingan',
          capaianKompetensi: 'Perlu Pendampingan',
          deskripsiTercapai: '',
          deskripsiPerluPeningkatan: '',
          deskripsiRapor: ''
        };
        resetMap[studentId] = cleared;
        clearedList.push(cleared);
      });
      return resetMap;
    });
    if (clearedList.length > 0) {
      bulkSaveNilai(clearedList, true);
    }
    showSaveNotification(`Nilai input untuk ${selectedMapel} berhasil dikosongkan!`);
  };

  const handleDownloadTemplate = () => {
    const schoolName = db.sekolah.namaSekolah || 'SD NEGERI BALIL';
    const schoolYear = db.sekolah.tahunPelajaran || '2026/2027';
    const pemda = db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG';
    const instansi = db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN';
    const localScores: AcademicScore[] = Object.values(rows);

    const wb = XLSX.utils.book_new();
    const columns = [
      'NISN',
      'Nama Siswa',
      'Kelas',
      'Semester',
      'Mata Pelajaran',
      'TP 1',
      'TP 2',
      'TP 3',
      'TP 4',
      'TP 5',
      'TP 6',
      'TP 7',
      'TP 8',
      'LM 1',
      'LM 2',
      'LM 3',
      'LM 4',
      'Non Tes',
      'Tes',
      'NA Sumatif TP',
      'NA Sumatif LM',
      'NA Sumatif SAS',
      'Nilai Akhir (Rapor)',
      'Deskripsi Capaian Tertinggi',
      'Deskripsi Capaian Terendah'
    ];

    const aoaData: (string | number)[][] = [
      [pemda],
      [instansi],
      [schoolName],
      [`REKAP NILAI KURIKULUM MERDEKA: ${selectedMapel.toUpperCase()}`],
      [`Kelas: ${selectedKelas} • Semester: ${selectedSemester} • Tahun Pelajaran: ${schoolYear}`],
      [],
      ['NO', ...columns],
      ...localScores.map((s, idx) => [
        idx + 1,
        s.nisn,
        s.namaSiswa,
        selectedKelas,
        selectedSemester,
        selectedMapel,
        s.tpScores?.tp1 ?? '',
        s.tpScores?.tp2 ?? '',
        s.tpScores?.tp3 ?? '',
        s.tpScores?.tp4 ?? '',
        s.tpScores?.tp5 ?? '',
        s.tpScores?.tp6 ?? '',
        s.tpScores?.tp7 ?? '',
        s.tpScores?.tp8 ?? '',
        s.lmScores?.lm1 ?? '',
        s.lmScores?.lm2 ?? '',
        s.lmScores?.lm3 ?? '',
        s.lmScores?.lm4 ?? '',
        s.sasNonTes ?? '',
        s.sasTes ?? '',
        s.naSumatifTp !== undefined ? s.naSumatifTp.toFixed(2) : '',
        s.naSumatifLm !== undefined ? s.naSumatifLm.toFixed(2) : '',
        s.naSumatifSas !== undefined ? s.naSumatifSas.toFixed(2) : '',
        s.nilaiAkhir || '',
        s.deskripsiTercapai || '',
        s.deskripsiPerluPeningkatan || ''
      ])
    ];

    const ws = XLSX.utils.aoa_to_sheet(aoaData);
    ws['!cols'] = [
      { wch: 5 },
      { wch: 16 },
      { wch: 28 },
      { wch: 8 },
      { wch: 10 },
      { wch: 28 },
      { wch: 7 }, { wch: 7 }, { wch: 7 }, { wch: 7 },
      { wch: 7 }, { wch: 7 }, { wch: 7 }, { wch: 7 },
      { wch: 7 }, { wch: 7 }, { wch: 7 }, { wch: 7 },
      { wch: 9 }, { wch: 9 },
      { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
      { wch: 45 }, { wch: 45 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Rekap Nilai');
    const safeName = selectedMapel.toLowerCase().replace(/[^a-z0-9]/g, '-');
    XLSX.writeFile(wb, `template-rekap-nilai-${safeName}-kelas-${selectedKelas}.xlsx`);
    showSaveNotification(`Template Rekap Nilai ${selectedMapel} berhasil diunduh!`);
  };

  const openLmRename = (lmKey: string, idx: number) => {
    setEditingLmKey(lmKey);
    setEditingLmDraft(lmTitles[lmKey] || `LM ${idx + 1}`);
  };

  const commitLmRename = () => {
    if (!editingLmKey) return;
    const trimmed = editingLmDraft.trim();
    const nextTitles = { ...lmTitles };
    if (!trimmed) {
      delete nextTitles[editingLmKey];
    } else {
      nextTitles[editingLmKey] = trimmed;
    }
    saveLingkupMateriTitles(lmScopeKey, nextTitles);
    setEditingLmKey(null);
    showSaveNotification('Nama kolom Lingkup Materi diperbarui!');
  };

  const displayMapelTitle =
    agamaFilter && selectedMapel.toLowerCase().includes('agama')
      ? `Agama ${agamaFilter}`
      : selectedMapel;

  return (
    <div className="space-y-4">
      {/* Header Toolbar */}
      <div className="bg-white rounded-2xl p-4 sm:p-5 border border-slate-200 shadow-xs flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="p-2.5 bg-cyan-50 text-cyan-800 rounded-xl border border-cyan-200">
            <FileSpreadsheet size={24} />
          </div>
          <div>
            <h2 className="text-lg font-extrabold text-slate-900">
              Input Nilai Siswa (Sumatif TP, Lingkup Materi &amp; SAS)
            </h2>
            <p className="text-xs text-slate-500">
              Kelas {selectedKelas} · Semester {selectedSemester} · NA Komponen, Nilai Akhir Rapor &amp; Deskripsi Capaian Tertinggi/Terendah terhitung otomatis.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          <button
            onClick={handleRecalculateAll}
            className="w-8 h-8 inline-flex items-center justify-center bg-cyan-600 hover:bg-cyan-700 text-white rounded-lg shadow-xs cursor-pointer transition-colors"
            title="Hitung Ulang NA & Deskripsi TP"
            aria-label="Hitung Ulang Nilai"
          >
            <RefreshCw size={15} />
          </button>
          <button
            onClick={handleConnectAllStudentsToAllSubjects}
            className="w-8 h-8 inline-flex items-center justify-center bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-500 hover:to-teal-600 text-white rounded-lg transition-all shadow-xs cursor-pointer"
            title="Hubungkan Semua Siswa ke Seluruh Mapel"
            aria-label="Hubungkan Semua Siswa ke Seluruh Mapel"
          >
            <Zap size={15} className="text-amber-300 fill-amber-300" />
          </button>
          <button
            onClick={() => {
              const count = syncAllScoresWithTPDescriptions();
              showSaveNotification(`Berhasil memperbarui ${count} data deskripsi nilai dari TP ke semua rapor!`);
            }}
            className="w-8 h-8 inline-flex items-center justify-center bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-all shadow-xs cursor-pointer"
            title="Hubungkan TP ke Rapor"
            aria-label="Hubungkan TP ke Rapor"
          >
            <CheckCircle2 size={15} />
          </button>
          <button
            onClick={() => openReportPreview()}
            className="w-8 h-8 inline-flex items-center justify-center bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-all shadow-xs cursor-pointer"
            title="Preview Rapor Lengkap Siswa"
            aria-label="Preview Rapor"
          >
            <Eye size={15} />
          </button>
          <div className="flex items-center gap-1 p-0.5 bg-slate-100 rounded-lg border border-slate-200">
            <button
              onClick={handleExportExcel}
              className="w-7 h-7 inline-flex items-center justify-center bg-emerald-600 hover:bg-emerald-700 text-white rounded-md transition-all shadow-xs cursor-pointer"
              title={`Ekspor Nilai ${selectedMapel} ke Excel (.xlsx)`}
              aria-label="Ekspor Mapel Excel"
            >
              <FileSpreadsheet size={14} />
            </button>
            <button
              onClick={handleExportCsv}
              className="w-7 h-7 inline-flex items-center justify-center bg-slate-700 hover:bg-slate-800 text-white rounded-md transition-all shadow-xs cursor-pointer"
              title={`Ekspor Nilai ${selectedMapel} ke CSV (.csv)`}
              aria-label="Ekspor CSV"
            >
              <Download size={14} />
            </button>
            <button
              onClick={handleExportFullWorkbook}
              className="w-7 h-7 inline-flex items-center justify-center bg-amber-500 hover:bg-amber-600 text-slate-950 rounded-md transition-all shadow-xs cursor-pointer"
              title="Ekspor Buku Nilai Sekelas (.xlsx Multi-Sheet)"
              aria-label="Ekspor Buku Nilai Sekelas"
            >
              <FileSpreadsheet size={14} />
            </button>
          </div>
          <button
            onClick={() => setActiveMenu('template_data')}
            className="w-8 h-8 inline-flex items-center justify-center bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-lg transition-colors cursor-pointer border border-indigo-200"
            title="Unduh Template Nilai"
            aria-label="Unduh Template Nilai"
          >
            <FolderDown size={15} />
          </button>
          <div className="flex items-center gap-1.5">
            <button
              onClick={handleSaveAll}
              disabled={saveStatus === 'saving'}
              title="Simpan Semua Nilai Mapel Ini"
              aria-label="Simpan Semua Nilai Mapel Ini"
              className="w-8 h-8 inline-flex items-center justify-center bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg transition-all shadow-xs cursor-pointer disabled:opacity-60"
            >
              <Save size={15} />
            </button>
            <SaveStatusIndicator status={saveStatus} compact lastSavedText={lastSavedTime} />
          </div>
        </div>
      </div>

      {/* Standard Action Toolbar */}
      <StandardActionToolbar
        title="Aksi Rekap Nilai"
        badge={displayMapelTitle}
        onSave={handleSaveAll}
        saveLabel="Simpan Nilai"
        saveStatus={saveStatus}
        onUpdate={handleRecalculateAll}
        updateLabel="Hitung Ulang NA"
        onEdit={() => setShowHiddenColumns((prev) => !prev)}
        editLabel={showHiddenColumns ? 'Hide Empty Columns' : 'Show Hidden Column'}
        isEditing={showHiddenColumns}
        onCancel={() => handleMapelChange(selectedMapel)}
        cancelLabel="Muat Ulang"
        onDelete={handleClearScores}
        deleteLabel="Kosongkan Nilai"
        onImport={handleImportFile}
        importLabel="Impor Excel"
        onExport={handleExportExcel}
        exportLabel="Expor Excel (.xlsx)"
        onUpload={handleImportFile}
        uploadLabel="Upload File Nilai"
        onClear={handleClearScores}
        clearLabel="Bersihkan Nilai"
        onTemplate={handleDownloadTemplate}
        templateLabel="Template Rekap Nilai"
      />

      {/* EXCEL-STYLE REKAP NILAI TOP BANNER */}
      <div className="rounded-xl border-2 border-slate-800 overflow-hidden shadow-sm">
        <div className="bg-cyan-300 px-3 py-2.5 flex flex-wrap items-center justify-between gap-3 border-b-2 border-slate-800">
          {/* Kiri: Cluster Tombol Agama (2x3) seperti Excel Rekap Nilai */}
          <div className="flex items-center gap-2.5">
            <div className="bg-slate-950 p-1.5 rounded-lg border-2 border-red-600 grid grid-cols-3 gap-1 shadow-xs">
              {RELIGION_TABS.map((tab) => {
                const isActive = agamaFilter.toLowerCase() === tab.key.toLowerCase();
                return (
                  <button
                    key={tab.key}
                    type="button"
                    onClick={() => handleReligionTabClick(tab.key)}
                    title={`Filter siswa agama ${tab.label} (klik lagi untuk menampilkan semua)`}
                    className={`px-2 py-1 text-[10px] font-black tracking-wide rounded-xs border transition-all cursor-pointer ${tab.bg} ${
                      isActive ? 'ring-2 ring-white scale-[1.03]' : 'opacity-90'
                    }`}
                  >
                    {tab.label}
                  </button>
                );
              })}
            </div>
            {agamaFilter && (
              <button
                type="button"
                onClick={() => setAgamaFilter('')}
                className="px-2 py-1 text-[10px] font-bold bg-white hover:bg-slate-100 text-slate-900 border border-slate-700 rounded-md cursor-pointer inline-flex items-center gap-1"
                title="Tampilkan Semua Agama"
              >
                <X size={11} />
                <span>Semua</span>
              </button>
            )}
          </div>

          {/* Tengah: Identitas REKAP NILAI & Pilihan Mapel + Rumus Bobot Aktif */}
          <div className="flex-1 min-w-[260px] text-slate-950">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-black tracking-[0.2em] uppercase">REKAP NILAI</span>
              <button
                type="button"
                onClick={() => setActiveMenu('kkm_pelajaran')}
                className="px-2 py-0.5 text-[10px] font-bold bg-white/90 hover:bg-white text-slate-900 border border-slate-700 rounded-sm cursor-pointer transition-colors"
                title="Ubah KKM, Bobot Nilai Rapor & KKTP"
              >
                {(() => {
                  const b = db.kkmSettings?.bobot;
                  const bTp = b?.namaPenilaianTp === 'FORMATIF' ? 0 : (b?.bobotTp ?? 1);
                  const bLm = b?.bobotLm ?? 1;
                  const bSas = b?.bobotSas ?? 2;
                  const tot = bTp + bLm + bSas || 1;
                  return bTp > 0
                    ? `Rumus NA = (${bTp}*S.TP + ${bLm}*S.LM + ${bSas}*S.AS)/${tot}`
                    : `Rumus NA = (${bLm}*S.LM + ${bSas}*S.AS)/${tot}`;
                })()}
              </button>
            </div>
            <div className="grid grid-cols-[175px_1fr] gap-x-2 text-xs font-bold leading-tight mt-0.5">
              <span>NAMA SEKOLAH</span>
              <span>: {db.sekolah.namaSekolah || 'SD Negeri Balil'}</span>
              <span>MUATAN/MATA PELAJARAN</span>
              <span className="flex items-center gap-2 flex-wrap">
                <span>: {displayMapelTitle}</span>
              </span>
            </div>
          </div>

          {/* Kanan: Pilih Mapel, Filter Siswa & Tombol Show Hidden Column */}
          <div className="flex flex-wrap items-center gap-2">
            <div className="inline-flex items-center gap-1.5 bg-white/95 border border-slate-700 rounded-lg px-2 py-1">
              <BookOpen size={14} className="text-cyan-800 shrink-0" />
              <select
                value={selectedMapel}
                onChange={(e) => handleMapelChange(e.target.value)}
                className="text-xs font-bold bg-transparent text-slate-900 focus:outline-hidden cursor-pointer max-w-[210px]"
              >
                {db.mataPelajaran.map((m) => (
                  <option key={m.kode || m.mataPelajaran} value={m.mataPelajaran}>
                    {m.mataPelajaran} ({m.kelompok})
                  </option>
                ))}
              </select>
            </div>

            <StudentFilterSelect compact className="bg-white/95 border-slate-700" />

            <button
              type="button"
              onClick={() => setShowHiddenColumns((prev) => !prev)}
              className={`px-3 py-1.5 rounded-lg text-xs font-extrabold border-2 transition-all cursor-pointer inline-flex items-center gap-1.5 shadow-xs ${
                showHiddenColumns
                  ? 'bg-slate-900 text-cyan-300 border-slate-900'
                  : 'bg-sky-200 hover:bg-sky-100 text-slate-950 border-slate-700'
              }`}
              title="Tampilkan atau sembunyikan kolom TP & Lingkup Materi tambahan"
            >
              <Columns3 size={14} />
              <span>{showHiddenColumns ? 'Hide Extra Column' : 'Show Hidden Column'}</span>
            </button>
          </div>
        </div>

        {/* Inline Edit Modal/Bar for Lingkup Materi Column Title */}
        {editingLmKey && (
          <div className="bg-sky-50 px-4 py-2 border-b border-slate-400 flex flex-wrap items-center gap-2 text-xs">
            <span className="font-bold text-slate-800">
              Ubah Nama Topik ({editingLmKey.toUpperCase()}):
            </span>
            <input
              type="text"
              value={editingLmDraft}
              onChange={(e) => setEditingLmDraft(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') commitLmRename();
                if (e.key === 'Escape') setEditingLmKey(null);
              }}
              placeholder="Contoh: Cerita Ramayana / Hyang Widhi Wasa"
              className="px-2.5 py-1 border border-sky-400 rounded-md bg-white text-slate-900 font-semibold min-w-[240px]"
              autoFocus
            />
            <button
              type="button"
              onClick={commitLmRename}
              className="px-2.5 py-1 bg-sky-700 hover:bg-sky-800 text-white font-bold rounded-md cursor-pointer"
            >
              Simpan Nama
            </button>
            <button
              type="button"
              onClick={() => setEditingLmKey(null)}
              className="px-2.5 py-1 bg-slate-200 hover:bg-slate-300 text-slate-800 font-bold rounded-md cursor-pointer"
            >
              Batal
            </button>
          </div>
        )}

        {/* SPREADSHEET TABLE */}
        <div className="overflow-x-auto bg-white">
          <table className="w-full border-collapse text-[11px]">
            <thead>
              {/* GROUP HEADER ROW 1 */}
              <tr className="text-center font-black uppercase tracking-tight">
                <th
                  rowSpan={2}
                  className="sticky left-0 z-20 bg-cyan-300 text-slate-950 border border-slate-700 px-2 py-2 w-9 min-w-[36px]"
                >
                  NO
                </th>
                <th
                  rowSpan={2}
                  className="sticky left-[36px] z-20 bg-cyan-300 text-slate-950 border border-slate-700 px-3 py-2 min-w-[190px] text-left"
                >
                  <div>NAMA SISWA</div>
                  <div className="text-[9px] font-semibold text-slate-700 normal-case mt-0.5">
                    {displayedSiswa.length} Siswa · Input Nilai 0–100
                  </div>
                </th>

                <th
                  colSpan={visibleTpKeys.length}
                  className="bg-lime-400 text-slate-950 border border-slate-700 py-1.5 px-2 text-xs font-black"
                >
                  {db.kkmSettings?.bobot?.namaPenilaianTp === 'FORMATIF'
                    ? 'FORMATIF (PERTIMBANGAN DESKRIPSI)'
                    : 'SUMATIF TUJUAN PEMBELAJARAN'}
                </th>

                <th
                  colSpan={visibleLmKeys.length}
                  className="bg-sky-400 text-slate-950 border border-slate-700 py-1.5 px-2 text-xs font-black"
                >
                  SUMATIF LINGKUP MATERI
                </th>

                <th
                  colSpan={2}
                  className="bg-cyan-300 text-slate-950 border border-slate-700 py-1.5 px-2 text-[10px] font-black leading-tight"
                >
                  SUMATIF AKHIR SEMESTER
                </th>

                <th
                  rowSpan={2}
                  className="bg-lime-400 text-slate-950 border border-slate-700 px-1.5 py-2 w-20 min-w-[78px] text-[9px] font-black leading-tight"
                >
                  NA SUMATIF TUJUAN PEMBELAJARAN
                </th>
                <th
                  rowSpan={2}
                  className="bg-sky-400 text-slate-950 border border-slate-700 px-1.5 py-2 w-20 min-w-[76px] text-[9px] font-black leading-tight"
                >
                  NA SUMATIF LINGKUP MATERI
                </th>
                <th
                  rowSpan={2}
                  className="bg-cyan-300 text-slate-950 border border-slate-700 px-1.5 py-2 w-20 min-w-[76px] text-[9px] font-black leading-tight"
                >
                  NA SUMATIF AKHIR SEMESTER
                </th>
                <th
                  rowSpan={2}
                  className="bg-pink-400 text-slate-950 border border-slate-700 px-2 py-2 w-20 min-w-[76px] text-[10px] font-black leading-tight"
                >
                  NILAI AKHIR (RAPOR)
                </th>

                <th
                  colSpan={3}
                  className="bg-slate-950 text-white border border-slate-700 py-1.5 px-3 text-xs font-black tracking-wide"
                >
                  DESKRIPSI RAPOR
                </th>
              </tr>

              {/* SUB-HEADER ROW 2 */}
              <tr className="text-center font-extrabold">
                {/* TP Columns */}
                {visibleTpKeys.map((tpKey, idx) => {
                  const tpMeta = activeTpSlots[idx];
                  const isExtraUnused = idx >= 8;
                  return (
                    <th
                      key={tpKey}
                      onClick={() => setActiveTpInfoIdx(activeTpInfoIdx === idx ? null : idx)}
                      title={
                        tpMeta
                          ? `TP ${idx + 1} (${tpMeta.kodeTp}): ${tpMeta.tujuanPembelajaran}`
                          : `TP ${idx + 1} (Klik untuk melihat detail atau input nilai 0-100)`
                      }
                      className={`border border-slate-700 px-1 py-1.5 w-12 min-w-[46px] h-16 align-middle cursor-pointer select-none ${
                        TP_HEADER_COLORS[idx % TP_HEADER_COLORS.length]
                      } ${isExtraUnused ? 'opacity-90' : ''}`}
                    >
                      <div className="flex flex-col items-center justify-center leading-tight">
                        <span className="text-[10px] font-black">TP {idx + 1}</span>
                        {tpMeta && (
                          <span className="text-[8px] font-semibold opacity-80 truncate max-w-[40px]">
                            {tpMeta.kodeTp}
                          </span>
                        )}
                      </div>
                    </th>
                  );
                })}

                {/* Lingkup Materi Columns */}
                {visibleLmKeys.map((lmKey, idx) => {
                  const customName = lmTitles[lmKey];
                  return (
                    <th
                      key={lmKey}
                      onClick={() => openLmRename(lmKey, idx)}
                      title={`${customName || `Lingkup Materi ${idx + 1}`} — Klik untuk ubah nama topik`}
                      className={`border border-slate-700 px-1 py-1.5 w-14 min-w-[52px] h-16 align-middle cursor-pointer group ${
                        LM_HEADER_COLORS[idx % LM_HEADER_COLORS.length]
                      }`}
                    >
                      <div className="flex flex-col items-center justify-center gap-0.5 leading-tight">
                        <span className="text-[9px] font-black line-clamp-2 max-w-[48px]">
                          {customName || `LM ${idx + 1}`}
                        </span>
                        <Edit3 size={9} className="opacity-50 group-hover:opacity-100" />
                      </div>
                    </th>
                  );
                })}

                {/* SAS Non Tes & Tes */}
                <th className="bg-lime-300 text-slate-950 border border-slate-700 px-1 py-1.5 w-13 min-w-[50px] text-[9px] font-black leading-tight">
                  NON TES
                </th>
                <th className="bg-teal-400 text-slate-950 border border-slate-700 px-1 py-1.5 w-13 min-w-[50px] text-[9px] font-black leading-tight">
                  TES
                </th>

                {/* Deskripsi Capaian Tertinggi, Terendah & Deskripsi Rapor */}
                <th className="bg-sky-600 text-white border border-slate-700 px-3 py-2 min-w-[240px] text-[10px] font-extrabold">
                  DESKRIPSI CAPAIAN TERTINGGI
                </th>
                <th className="bg-blue-700 text-white border border-slate-700 px-3 py-2 min-w-[240px] text-[10px] font-extrabold">
                  DESKRIPSI CAPAIAN TERENDAH
                </th>
                <th className="bg-indigo-800 text-white border border-slate-700 px-3 py-2 min-w-[280px] text-[10px] font-extrabold">
                  DESKRIPSI RAPOR (OTOMATIS)
                </th>
              </tr>
            </thead>

            <tbody>
              {displayedSiswa.map((s, idx) => {
                const item = rows[s.idSiswa] || buildInitialRow(s, selectedMapel);
                const matchedTps = getMatchedTpsForSubjectAndStudent(
                  selectedMapel,
                  selectedKelas,
                  db.tujuanPembelajaran,
                  s.agama
                );
                const matchedSlots = resolveTpSlots(matchedTps);
                const callName = formatStudentCallName(s.namaSiswa, s.bukuInduk?.namaPanggilan);
                const km = hasKurikulumMerdekaInputs(item)
                  ? calculateKurikulumMerdekaScores(
                      item,
                      matchedTps,
                      selectedMapel,
                      lmTitles,
                      db.kkmSettings,
                      callName
                    )
                  : null;

                const naTpDisplay = km?.naSumatifTp ?? item.naSumatifTp;
                const naLmDisplay = km?.naSumatifLm ?? item.naSumatifLm;
                const naSasDisplay = km?.naSumatifSas ?? item.naSumatifSas;
                const displayNisn = !isPlaceholderNisn(s.nisn) && !s.nisn.startsWith('BLM_') ? s.nisn : '-';

                return (
                  <tr key={s.idSiswa} className="hover:bg-amber-50/40 transition-colors">
                    <td className="sticky left-0 z-10 bg-cyan-200 text-slate-950 border border-slate-700 px-2 py-1.5 text-center font-mono font-bold">
                      {idx + 1}
                    </td>

                    <td className="sticky left-[36px] z-10 bg-cyan-100 text-slate-950 border border-slate-700 px-2.5 py-1.5">
                      <div className="font-bold text-slate-900 leading-snug">{s.namaSiswa}</div>
                      <div className="flex items-center gap-1.5 text-[9px] text-slate-600 font-mono">
                        <span>{displayNisn}</span>
                        {s.agama && (
                          <>
                            <span>·</span>
                            <span className="font-sans font-semibold text-cyan-900">{s.agama}</span>
                          </>
                        )}
                      </div>
                    </td>

                    {/* TP Inputs */}
                    {visibleTpKeys.map((tpKey, tpIdx) => {
                      const val = item.tpScores?.[tpKey];
                      const isExtraSlot = tpIdx >= 8 && val === undefined;
                      return (
                        <td
                          key={tpKey}
                          className={`border border-slate-700 p-0 text-center ${
                            isExtraSlot ? 'bg-red-600' : 'bg-white'
                          }`}
                        >
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={val ?? ''}
                            onChange={(e) => updateTpCell(s.idSiswa, tpKey, e.target.value)}
                            title={`Input Nilai 0-100 · TP ${tpIdx + 1}${matchedSlots[tpIdx] ? `: ${matchedSlots[tpIdx]?.tujuanPembelajaran}` : ''}`}
                            className={`w-12 h-9 px-0.5 text-center font-mono tabular-nums text-xs font-bold focus:outline-hidden focus:ring-2 focus:ring-inset focus:ring-blue-600 ${
                              isExtraSlot
                                ? 'bg-red-600 text-white placeholder-white/60 focus:bg-white focus:text-slate-950'
                                : 'bg-white text-slate-950'
                            }`}
                          />
                        </td>
                      );
                    })}

                    {/* Lingkup Materi Inputs */}
                    {visibleLmKeys.map((lmKey, lmIdx) => {
                      const val = item.lmScores?.[lmKey];
                      const isExtraSlot = lmIdx >= 4 && val === undefined;
                      return (
                        <td
                          key={lmKey}
                          className={`border border-slate-700 p-0 text-center ${
                            isExtraSlot ? 'bg-red-600' : 'bg-white'
                          }`}
                        >
                          <input
                            type="number"
                            min={0}
                            max={100}
                            value={val ?? ''}
                            onChange={(e) => updateLmCell(s.idSiswa, lmKey, e.target.value)}
                            title={`Input Nilai 0-100 · ${lmTitles[lmKey] || `LM ${lmIdx + 1}`}`}
                            className={`w-13 h-9 px-0.5 text-center font-mono tabular-nums text-xs font-bold focus:outline-hidden focus:ring-2 focus:ring-inset focus:ring-sky-600 ${
                              isExtraSlot
                                ? 'bg-red-600 text-white placeholder-white/60 focus:bg-white focus:text-slate-950'
                                : 'bg-white text-slate-950'
                            }`}
                          />
                        </td>
                      );
                    })}

                    {/* SAS Non Tes & Tes */}
                    <td className="border border-slate-700 p-0 text-center bg-white">
                      <input
                        type="number"
                        min={0}
                        max={100}
                        value={item.sasNonTes ?? ''}
                        onChange={(e) => updateSasCell(s.idSiswa, 'sasNonTes', e.target.value)}
                        title="Input Nilai 0-100 · SAS Non Tes"
                        className="w-13 h-9 px-0.5 text-center font-mono tabular-nums text-xs font-bold bg-white text-slate-950 focus:outline-hidden focus:ring-2 focus:ring-inset focus:ring-teal-600"
                      />
                    </td>
                    <td className="border border-slate-700 p-0 text-center bg-white">
                      <input
                        type="number"
                        min={0}
                        max={100}
                        value={item.sasTes ?? ''}
                        onChange={(e) => updateSasCell(s.idSiswa, 'sasTes', e.target.value)}
                        title="Input Nilai 0-100 · SAS Tes"
                        className="w-13 h-9 px-0.5 text-center font-mono tabular-nums text-xs font-bold bg-white text-slate-950 focus:outline-hidden focus:ring-2 focus:ring-inset focus:ring-teal-600"
                      />
                    </td>

                    {/* NA Sumatif TP, LM, SAS */}
                    <td className="border border-slate-700 bg-cyan-200/85 text-slate-950 px-1.5 py-1 text-center font-mono tabular-nums font-bold">
                      {naTpDisplay !== undefined ? naTpDisplay.toFixed(2) : ''}
                    </td>
                    <td className="border border-slate-700 bg-cyan-200/85 text-slate-950 px-1.5 py-1 text-center font-mono tabular-nums font-bold">
                      {naLmDisplay !== undefined ? naLmDisplay.toFixed(2) : ''}
                    </td>
                    <td className="border border-slate-700 bg-cyan-200/85 text-slate-950 px-1.5 py-1 text-center font-mono tabular-nums font-bold">
                      {naSasDisplay !== undefined ? naSasDisplay.toFixed(2) : ''}
                    </td>

                    {/* NILAI AKHIR (RAPOR) */}
                    <td className="border border-slate-700 p-0 text-center bg-cyan-200">
                      <input
                        type="number"
                        min={0}
                        max={100}
                        value={item.nilaiAkhir > 0 ? item.nilaiAkhir : ''}
                        onChange={(e) => updateManualField(s.idSiswa, 'nilaiAkhir', e.target.value)}
                        title="Nilai Akhir Rapor (otomatis rata-rata NA TP, NA LM & NA SAS, dapat disesuaikan manual)"
                        className="w-16 h-9 px-1 text-center font-mono tabular-nums text-sm font-black italic bg-cyan-200 focus:bg-white text-slate-950 focus:outline-hidden focus:ring-2 focus:ring-inset focus:ring-pink-500"
                      />
                    </td>

                    {/* DESKRIPSI CAPAIAN TERTINGGI */}
                    <td className="border border-slate-700 p-1 bg-cyan-100/70">
                      <textarea
                        rows={2}
                        placeholder="Deskripsi capaian tertinggi otomatis dari nilai TP tertinggi..."
                        value={item.deskripsiTercapai || ''}
                        onChange={(e) => updateManualField(s.idSiswa, 'deskripsiTercapai', e.target.value)}
                        className="w-full min-w-[240px] px-1.5 py-1 text-[11px] bg-white/85 focus:bg-white border border-cyan-300 rounded-sm leading-snug text-slate-900 focus:outline-hidden focus:ring-1 focus:ring-sky-600"
                      />
                    </td>

                    {/* DESKRIPSI CAPAIAN TERENDAH */}
                    <td className="border border-slate-700 p-1 bg-cyan-100/70">
                      <textarea
                        rows={2}
                        placeholder="Deskripsi capaian terendah otomatis dari nilai TP terendah..."
                        value={item.deskripsiPerluPeningkatan || ''}
                        onChange={(e) => updateManualField(s.idSiswa, 'deskripsiPerluPeningkatan', e.target.value)}
                        className="w-full min-w-[220px] px-1.5 py-1 text-[11px] bg-white/85 focus:bg-white border border-cyan-300 rounded-sm leading-snug text-slate-900 focus:outline-hidden focus:ring-1 focus:ring-blue-700"
                      />
                    </td>

                    {/* DESKRIPSI RAPOR (OTOMATIS GABUNGAN TERTINGGI + TERENDAH) */}
                    <td className="border border-slate-700 p-1 bg-indigo-50/80">
                      <textarea
                        rows={2}
                        placeholder="Deskripsi rapor otomatis dari gabungan capaian tertinggi & terendah..."
                        value={item.deskripsiRapor || ''}
                        onChange={(e) => updateManualField(s.idSiswa, 'deskripsiRapor', e.target.value)}
                        className="w-full min-w-[260px] px-1.5 py-1 text-[11px] bg-white/90 focus:bg-white border border-indigo-300 rounded-sm leading-snug text-slate-900 focus:outline-hidden focus:ring-1 focus:ring-indigo-700"
                      />
                    </td>
                  </tr>
                );
              })}

              {displayedSiswa.length === 0 && (
                <tr>
                  <td
                    colSpan={9 + visibleTpKeys.length + visibleLmKeys.length}
                    className="py-8 text-center text-xs text-slate-500 font-medium border border-slate-300"
                  >
                    Tidak ada siswa yang sesuai dengan filter kelas/agama/nama saat ini.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Referensi Tujuan Pembelajaran (TP 1 .. TP 12) untuk Mapel & Kelas Aktif */}
      <div className="bg-white rounded-xl p-4 border border-slate-200 shadow-xs space-y-2">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="text-xs font-bold text-slate-800">
            Referensi Kolom Tujuan Pembelajaran ({displayMapelTitle} · Kelas {selectedKelas})
          </div>
          <button
            type="button"
            onClick={() => setActiveMenu('tujuan_pembelajaran')}
            className="text-xs font-bold text-cyan-800 hover:text-cyan-950 underline cursor-pointer"
          >
            Kelola Daftar Tujuan Pembelajaran (TP)
          </button>
        </div>

        {activeTpSlots.some(Boolean) ? (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-2 text-[11px]">
            {activeTpSlots.map((tp, idx) =>
              tp ? (
                <div
                  key={tp.id}
                  className={`p-2 rounded-lg border transition-colors ${
                    activeTpInfoIdx === idx
                      ? 'bg-cyan-50 border-cyan-500 text-slate-950'
                      : 'bg-slate-50 border-slate-200 text-slate-700'
                  }`}
                >
                  <div className="font-extrabold text-slate-900">
                    TP {idx + 1} · {tp.kodeTp}
                  </div>
                  <div className="line-clamp-2 mt-0.5 leading-snug">{tp.tujuanPembelajaran}</div>
                </div>
              ) : null
            )}
          </div>
        ) : (
          <p className="text-xs text-slate-500">
            Belum ada TP khusus untuk {displayMapelTitle} Kelas {selectedKelas}. Anda tetap dapat mengisi nilai TP 1–TP 12 secara bebas; deskripsi rapor akan menggunakan frasa standar mata pelajaran.
          </p>
        )}
      </div>
    </div>
  );
};
