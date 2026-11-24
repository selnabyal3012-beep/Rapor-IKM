import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  BookUser,
  Printer,
  Download,
  FileSpreadsheet,
  Edit3,
  ChevronLeft,
  ChevronRight,
  Users,
  User,
  FileText,
  BookOpen,
  CheckSquare,
  Square,
  Search,
  X,
  Save,
  Upload,
  Eye,
  EyeOff,
  FileCheck2,
  Zap,
  GraduationCap,
  Maximize2,
  Minimize2
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { useApp, PAPER_SIZE_MAP, getReportFontCssStack } from '../../context/AppContext';
import {
  BukuIndukProfile,
  ClassLevel,
  SD_CLASSES,
  SemesterType,
  Student,
  getFaseFromKelas
} from '../../types';
import { LogoTutWuriHandayani } from '../common/OfficialLogos';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { PrintTypographyControl } from '../common/PrintTypographyControl';
import { matchesStudentRecord, matchesStudentScore, isPlaceholderNisn } from '../../utils/studentSyncHelper';
import { generateReportDescriptionFromTP, getNamaPelajaranAgama } from '../../utils/tpDescriptionHelper';

type BukuIndukPageTab = 'semua' | 'kaver' | 'petunjuk' | 'biodata' | 'nilai';

function formatDateIndo(dateStr?: string): string {
  if (!dateStr || dateStr === '-' || dateStr === '_') return '-';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    const months = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
  } catch {
    return dateStr;
  }
}

function getPaperCssSize(pSize?: string, orientation: 'portrait' | 'landscape' = 'portrait'): string {
  const key = (pSize as keyof typeof PAPER_SIZE_MAP) || 'F4';
  const spec = PAPER_SIZE_MAP[key] || PAPER_SIZE_MAP.F4;
  return orientation === 'landscape'
    ? `${spec.heightMm}mm ${spec.widthMm}mm`
    : `${spec.widthMm}mm ${spec.heightMm}mm`;
}

function cleanVal(val?: string | null, fallback = '-'): string {
  if (!val) return fallback;
  const trimmed = val.trim();
  if (!trimmed || trimmed === '_' || trimmed === '-') return fallback;
  return trimmed;
}

function displayNisn(nisn?: string): string {
  if (!nisn || isPlaceholderNisn(nisn) || nisn.startsWith('BLM_')) return '-';
  return nisn;
}

export const BukuIndukSection: React.FC = () => {
  const {
    db,
    selectedKelas,
    setSelectedKelas,
    selectedSemester,
    setSelectedSemester,
    filteredStudent,
    printTargetStudentId,
    setPrintTargetStudentId,
    saveSiswa,
    syncAllScoresWithTPDescriptions,
    showSaveNotification,
    printFontFamily,
    printFontSize,
    setPrintFontSize,
    printLineHeight,
    printTextAlign,
    paperSize,
    setPaperSize,
    paperOrientation,
    setPaperOrientation,
    paperMargins,
    printScale
  } = useApp();

  const [activeTab, setActiveTab] = useState<BukuIndukPageTab>('semua');
  const [printScope, setPrintScope] = useState<'single' | 'batch'>('single');
  const [searchStudent, setSearchStudent] = useState('');
  const [isPreviewVisible, setIsPreviewVisible] = useState(true);
  const [filterPanelMode, setFilterPanelMode] = useState<'full' | 'compact'>('compact');
  const [selectedStudentIds, setSelectedStudentIds] = useState<string[]>([]);
  const [transcriptSemester, setTranscriptSemester] = useState<SemesterType>(selectedSemester);

  // Edit Modal State
  const [editModalOpen, setEditModalOpen] = useState(false);
  const [editStudentTarget, setEditStudentTarget] = useState<Student | null>(null);
  const [formBase, setFormBase] = useState<Partial<Student>>({});
  const [formBi, setFormBi] = useState<BukuIndukProfile>({});
  const [activeEditSection, setActiveEditSection] = useState<'pribadi' | 'kesehatan' | 'ortu' | 'perkembangan'>('pribadi');
  const photoInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setTranscriptSemester(selectedSemester);
  }, [selectedSemester]);

  useEffect(() => {
    if (
      filteredStudent &&
      filteredStudent.kelas === selectedKelas &&
      printTargetStudentId !== filteredStudent.idSiswa
    ) {
      setPrintTargetStudentId(filteredStudent.idSiswa);
    }
  }, [filteredStudent, selectedKelas, printTargetStudentId, setPrintTargetStudentId]);

  const siswaKelas = useMemo(() => {
    return db.siswa.filter((s) => s.kelas === selectedKelas);
  }, [db.siswa, selectedKelas]);

  const filteredStudents = useMemo(() => {
    const base =
      filteredStudent && filteredStudent.kelas === selectedKelas
        ? siswaKelas.filter((s) => s.idSiswa === filteredStudent.idSiswa)
        : siswaKelas;
    if (!searchStudent.trim()) return base;
    const q = searchStudent.toLowerCase();
    return base.filter(
      (s) =>
        s.namaSiswa.toLowerCase().includes(q) ||
        (s.nisn || '').toLowerCase().includes(q) ||
        (s.nis || '').toLowerCase().includes(q)
    );
  }, [siswaKelas, filteredStudent, selectedKelas, searchStudent]);

  const currentStudent =
    filteredStudents.find((s) => s.idSiswa === printTargetStudentId) ||
    filteredStudents[0] ||
    siswaKelas[0];

  useEffect(() => {
    if (selectedStudentIds.length === 0 && siswaKelas.length > 0) {
      setSelectedStudentIds(siswaKelas.map((s) => s.idSiswa));
    }
  }, [siswaKelas, selectedStudentIds.length]);

  const studentsToRender = useMemo(() => {
    if (printScope === 'single') {
      return currentStudent ? [currentStudent] : [];
    }
    const chosen = filteredStudents.filter((s) => selectedStudentIds.includes(s.idSiswa));
    return chosen.length > 0 ? chosen : currentStudent ? [currentStudent] : [];
  }, [printScope, currentStudent, filteredStudents, selectedStudentIds]);

  const finalPageCssSize = getPaperCssSize(paperSize, paperOrientation);
  const reportFontCss = getReportFontCssStack(printFontFamily || 'Times New Roman');
  const sheetMaxWClass =
    paperOrientation === 'landscape'
      ? 'max-w-6xl report-paper-sheet--landscape'
      : 'max-w-4xl report-paper-sheet--portrait';

  const adjustFontSize = (direction: 'up' | 'down') => {
    const cur = parseFloat(printFontSize) || 11;
    const next =
      direction === 'up'
        ? Math.min(24, Math.round((cur + (cur >= 16 ? 2 : cur >= 12 ? 1 : 0.5)) * 2) / 2)
        : Math.max(8, Math.round((cur - (cur > 16 ? 2 : cur > 12 ? 1 : 0.5)) * 2) / 2);
    const nextStr = `${next}pt`;
    setPrintFontSize(nextStr);
    showSaveNotification(`Ukuran huruf Buku Induk diubah ke ${nextStr}`);
  };

  const handlePrevStudent = () => {
    if (!currentStudent || filteredStudents.length === 0) return;
    const idx = filteredStudents.findIndex((s) => s.idSiswa === currentStudent.idSiswa);
    if (idx > 0) {
      setPrintTargetStudentId(filteredStudents[idx - 1].idSiswa);
    }
  };

  const handleNextStudent = () => {
    if (!currentStudent || filteredStudents.length === 0) return;
    const idx = filteredStudents.findIndex((s) => s.idSiswa === currentStudent.idSiswa);
    if (idx >= 0 && idx < filteredStudents.length - 1) {
      setPrintTargetStudentId(filteredStudents[idx + 1].idSiswa);
    }
  };

  const toggleBatchStudent = (id: string) => {
    setSelectedStudentIds((prev) =>
      prev.includes(id) ? prev.filter((item) => item !== id) : [...prev, id]
    );
  };

  // Open Edit Modal for Buku Induk
  const handleOpenEditModal = (student: Student) => {
    setEditStudentTarget(student);
    setFormBase({
      namaSiswa: student.namaSiswa || '',
      nis: student.nis === '-' ? '' : student.nis || '',
      nisn: isPlaceholderNisn(student.nisn) || student.nisn.startsWith('BLM_') ? '' : student.nisn,
      jenisKelamin: student.jenisKelamin,
      tempatLahir: cleanVal(student.tempatLahir, ''),
      tanggalLahir: cleanVal(student.tanggalLahir, ''),
      agama: student.agama || 'Kristen',
      alamat: cleanVal(student.alamat, ''),
      telepon: cleanVal(student.telepon, ''),
      sekolahAsal: cleanVal(student.sekolahAsal, ''),
      diterimaKelas: cleanVal(student.diterimaKelas, `Kelas ${student.kelas}`),
      diterimaTanggal: cleanVal(student.diterimaTanggal, ''),
      namaAyah: cleanVal(student.namaAyah, ''),
      pekerjaanAyah: cleanVal(student.pekerjaanAyah, ''),
      namaIbu: cleanVal(student.namaIbu, ''),
      pekerjaanIbu: cleanVal(student.pekerjaanIbu, ''),
      namaWali: cleanVal(student.namaWali, ''),
      pekerjaanWali: cleanVal(student.pekerjaanWali, ''),
      alamatWali: cleanVal(student.alamatWali, ''),
      fotoUrl: student.fotoUrl || ''
    });
    setFormBi({
      kewarganegaraan: 'Indonesia',
      kewarganegaraanAyah: 'Indonesia',
      kewarganegaraanIbu: 'Indonesia',
      kewarganegaraanWali: 'Indonesia',
      statusAyah: 'Masih Hidup',
      statusIbu: 'Masih Hidup',
      anakYatim: 'Lengkap',
      tinggalDengan: 'Orang Tua',
      ...(student.bukuInduk || {})
    });
    setActiveEditSection('pribadi');
    setEditModalOpen(true);
  };

  const handlePhotoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) {
      showSaveNotification('Ukuran foto maksimal 2 MB agar penyimpanan tetap ringan.');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string') {
        setFormBase((prev) => ({ ...prev, fotoUrl: reader.result as string }));
      }
    };
    reader.readAsDataURL(file);
  };

  const handleSaveBukuInduk = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editStudentTarget) return;
    const rawNisn = (formBase.nisn || '').trim();
    const safeNisn =
      rawNisn === '' || rawNisn === '-' ? `BLM_${editStudentTarget.idSiswa}` : rawNisn;

    const updatedStudent: Student = {
      ...editStudentTarget,
      namaSiswa: (formBase.namaSiswa || editStudentTarget.namaSiswa).trim(),
      namaPanggilan: (formBi.namaPanggilan || editStudentTarget.namaPanggilan || '').trim(),
      nis: (formBase.nis || '').trim() || '-',
      nisn: safeNisn,
      jenisKelamin: (formBase.jenisKelamin as 'L' | 'P') || editStudentTarget.jenisKelamin,
      tempatLahir: (formBase.tempatLahir || '').trim() || '-',
      tanggalLahir: formBase.tanggalLahir || '-',
      agama: formBase.agama || editStudentTarget.agama,
      alamat: (formBase.alamat || '').trim() || '-',
      telepon: (formBase.telepon || '').trim(),
      sekolahAsal: (formBase.sekolahAsal || '').trim(),
      diterimaKelas: (formBase.diterimaKelas || '').trim(),
      diterimaTanggal: (formBase.diterimaTanggal || '').trim(),
      namaAyah: (formBase.namaAyah || '').trim(),
      pekerjaanAyah: (formBase.pekerjaanAyah || '').trim(),
      namaIbu: (formBase.namaIbu || '').trim(),
      pekerjaanIbu: (formBase.pekerjaanIbu || '').trim(),
      namaWali: (formBase.namaWali || '').trim(),
      pekerjaanWali: (formBase.pekerjaanWali || '').trim(),
      alamatWali: (formBase.alamatWali || '').trim(),
      fotoUrl: formBase.fotoUrl || '',
      bukuInduk: {
        ...formBi
      }
    };

    saveSiswa(updatedStudent, editStudentTarget.nisn);
    setEditModalOpen(false);
    showSaveNotification(`Biodata Buku Induk ${updatedStudent.namaSiswa} berhasil disimpan!`);
  };

  // Print & Export HTML/PDF
  const handlePrintBukuInduk = () => {
    setIsPreviewVisible(true);
    setTimeout(() => {
      window.print();
    }, 180);
  };

  const handleExportHtmlPdf = () => {
    const printArea = document.getElementById('buku-induk-print-area');
    if (!printArea) return;

    const label =
      studentsToRender.length === 1
        ? studentsToRender[0].namaSiswa
        : `Kelas_${selectedKelas}_${studentsToRender.length}_Siswa`;

    const htmlContent = `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <title>Buku Induk Siswa - ${label}</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    @page {
      size: ${finalPageCssSize};
      margin: ${paperMargins.top}mm ${paperMargins.right}mm ${paperMargins.bottom}mm ${paperMargins.left}mm;
    }
    body {
      font-family: ${reportFontCss};
      font-size: ${printFontSize};
      line-height: ${printLineHeight};
      color: #000;
      background: #fff;
      margin: 0;
      padding: 12px;
    }
    .report-typography-scope, .report-typography-scope * {
      font-family: ${reportFontCss} !important;
      line-height: ${printLineHeight};
    }
    .print-page-break {
      page-break-after: always;
      break-after: page;
      margin-bottom: 24px;
    }
    @media print {
      body { padding: 0; margin: 0; }
      .printable-area { zoom: ${printScale / 100}; }
      .no-print { display: none !important; }
    }
  </style>
</head>
<body>
  ${printArea.innerHTML}
</body>
</html>`;

    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `buku-induk-${label.toLowerCase().replace(/[^a-z0-9]+/gi, '-')}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showSaveNotification(`Dokumen Buku Induk (${studentsToRender.length} siswa) berhasil diunduh!`);
  };

  // Export Excel Buku Induk
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const targetList = studentsToRender.length > 0 ? studentsToRender : siswaKelas;

    const biodataRows = targetList.map((s, idx) => {
      const bi = s.bukuInduk || {};
      return {
        No: idx + 1,
        'No Induk (NIS)': cleanVal(s.nis),
        NISN: displayNisn(s.nisn),
        'Nama Lengkap': s.namaSiswa,
        'Nama Panggilan': cleanVal(bi.namaPanggilan),
        'L/P': s.jenisKelamin,
        'Tempat Lahir': cleanVal(s.tempatLahir),
        'Tanggal Lahir': formatDateIndo(s.tanggalLahir),
        Agama: cleanVal(s.agama),
        Kewarganegaraan: cleanVal(bi.kewarganegaraan, 'Indonesia'),
        'Anak Ke': cleanVal(bi.anakKe),
        'Saudara Kandung': cleanVal(bi.jumlahSaudaraKandung),
        'Bahasa Sehari-hari': cleanVal(bi.bahasaSehariHari),
        'Alamat Siswa': cleanVal(s.alamat),
        'No. Telepon': cleanVal(s.telepon),
        'Tinggal Dengan': cleanVal(bi.tinggalDengan),
        'Golongan Darah': cleanVal(bi.golonganDarah),
        'Tinggi/Berat Badan': [bi.tinggiBadan ? `${bi.tinggiBadan} cm` : '', bi.beratBadan ? `${bi.beratBadan} kg` : ''].filter(Boolean).join(' / ') || '-',
        'Asal Sekolah / TK': cleanVal(bi.asalTkPaud || s.sekolahAsal),
        'Diterima Kelas / Tanggal': `${cleanVal(s.diterimaKelas, `Kelas ${s.kelas}`)} (${cleanVal(s.diterimaTanggal)})`,
        'Nama Ayah': cleanVal(s.namaAyah),
        'Pekerjaan Ayah': cleanVal(s.pekerjaanAyah),
        'Pendidikan Ayah': cleanVal(bi.pendidikanAyah),
        'Nama Ibu': cleanVal(s.namaIbu),
        'Pekerjaan Ibu': cleanVal(s.pekerjaanIbu),
        'Pendidikan Ibu': cleanVal(bi.pendidikanIbu),
        'Nama Wali': cleanVal(s.namaWali),
        'Pekerjaan Wali': cleanVal(s.pekerjaanWali),
        'Kegemaran / Hobi': [bi.hobiKesenian, bi.hobiOlahraga, bi.hobiOrganisasi].filter(Boolean).join(', ') || '-',
        'Status / Kelulusan': cleanVal(bi.melanjutkanKe || s.statusSiswa)
      };
    });

    const wsBiodata = XLSX.utils.json_to_sheet(biodataRows);
    XLSX.utils.book_append_sheet(wb, wsBiodata, 'Buku Induk Biodata');

    const nilaiRows: Record<string, string | number>[] = [];
    targetList.forEach((s) => {
      db.mataPelajaran.forEach((m, mIdx) => {
        const namaMapel = getNamaPelajaranAgama(m.mataPelajaran, s.agama);
        const score = db.nilaiAkademik.find(
          (n) =>
            matchesStudentScore(n, s) &&
            n.mataPelajaran === m.mataPelajaran &&
            n.kelas === s.kelas &&
            n.semester === transcriptSemester
        );
        nilaiRows.push({
          'No Induk (NIS)': cleanVal(s.nis),
          NISN: displayNisn(s.nisn),
          'Nama Siswa': s.namaSiswa,
          Kelas: s.kelas,
          Semester: transcriptSemester,
          No: mIdx + 1,
          'Mata Pelajaran': namaMapel,
          Kelompok: m.kelompok,
          KKM: score?.kkm ?? m.kkm ?? 70,
          'Nilai Akhir': score?.nilaiAkhir ?? 0,
          Predikat: score?.predikat ?? '-',
          'Deskripsi Capaian': score?.deskripsiRapor || score?.deskripsiTercapai || '-'
        });
      });
    });

    const wsNilai = XLSX.utils.json_to_sheet(nilaiRows);
    XLSX.utils.book_append_sheet(wb, wsNilai, 'Transkrip Nilai');

    XLSX.writeFile(wb, `Buku_Induk_Siswa_Kelas_${selectedKelas}_${transcriptSemester}.xlsx`);
    showSaveNotification(`Buku Induk Siswa (${targetList.length} siswa) berhasil diekspor ke Excel (.xlsx)!`);
  };

  // Helper to resolve Wali Kelas info per class
  const getWaliForKelas = (kelas: ClassLevel) => {
    const kelasObj = db.kelas.find((k) => k.namaKelas === kelas);
    const namaGuru =
      (db.sekolah.kelas === kelas && db.sekolah.waliKelas) ||
      kelasObj?.waliKelas ||
      db.sekolah.waliKelas ||
      '-';
    const nip =
      (db.sekolah.kelas === kelas && db.sekolah.nipWaliKelas) ||
      db.sekolah.nipWaliKelas ||
      '-';
    return { namaGuru, nip };
  };

  // Render helper for a student's 4-page Buku Induk
  const renderStudentBukuInduk = (student: Student, studentIdx: number) => {
    const bi = student.bukuInduk || {};
    const nisDisplay = cleanVal(student.nis);
    const nisnDisplay = displayNisn(student.nisn);
    const wali = getWaliForKelas(student.kelas);
    const isSem2 = transcriptSemester === 'Genap' || transcriptSemester === '2';

    const showKaver = activeTab === 'semua' || activeTab === 'kaver';
    const showPetunjuk = activeTab === 'semua' || activeTab === 'petunjuk';
    const showBiodata = activeTab === 'semua' || activeTab === 'biodata';
    const showNilai = activeTab === 'semua' || activeTab === 'nilai';

    // Student transcript scores
    const studentScores = db.mataPelajaran.map((mapel) => {
      const found = db.nilaiAkademik.find(
        (n) =>
          matchesStudentScore(n, student) &&
          n.mataPelajaran === mapel.mataPelajaran &&
          n.kelas === student.kelas &&
          n.semester === transcriptSemester
      );
      const autoDesc =
        found && found.nilaiAkhir > 0
          ? generateReportDescriptionFromTP(
              mapel.mataPelajaran,
              student.kelas,
              found.nilaiAkhir,
              db.tujuanPembelajaran,
              student.namaSiswa,
              student.agama
            ).deskripsiRapor
          : '';

      const desc =
        found?.deskripsiRapor && found.deskripsiRapor.trim() !== '' && found.deskripsiRapor !== '-'
          ? found.deskripsiRapor
          : autoDesc || found?.deskripsiTercapai || '-';

      return {
        mapel,
        namaTampil: getNamaPelajaranAgama(mapel.mataPelajaran, student.agama),
        kkm: found?.kkm ?? mapel.kkm ?? 70,
        nilaiAkhir: found?.nilaiAkhir ?? 0,
        capaian: found?.capaianKompetensi || found?.predikat || '-',
        deskripsi: desc
      };
    });

    const wajibScores = studentScores.filter((s) => s.mapel.kelompok === 'Wajib');
    const pilihanMulokScores = studentScores.filter((s) => s.mapel.kelompok !== 'Wajib');
    const scoredItems = studentScores.filter((s) => s.nilaiAkhir > 0);
    const totalNilai = scoredItems.reduce((acc, s) => acc + s.nilaiAkhir, 0);
    const rataRata = scoredItems.length > 0 ? (totalNilai / scoredItems.length).toFixed(2) : '-';

    const studentEkskul = db.ekstrakurikuler.filter(
      (e) =>
        matchesStudentRecord(e, student) &&
        e.kelas === student.kelas &&
        e.semester === transcriptSemester
    );

    const studentAbsen = db.kehadiran.find(
      (k) =>
        matchesStudentRecord(k, student) &&
        k.kelas === student.kelas &&
        k.semester === transcriptSemester
    ) || { sakit: 0, izin: 0, alpa: 0 };

    const studentNote = db.catatanWali.find(
      (c) =>
        matchesStudentRecord(c, student) &&
        c.kelas === student.kelas &&
        c.semester === transcriptSemester
    );

    const studentDecision = db.keputusanSiswa?.find(
      (d) =>
        matchesStudentRecord(d, student) &&
        d.kelas === student.kelas &&
        d.semester === transcriptSemester
    );

    const placeAndDate =
      [db.sekolah.tempatRapor, db.sekolah.tanggalRapor]
        .filter((v) => v && v.trim() !== '' && v !== '-')
        .join(', ') || 'Oksibil, ..............................';

    return (
      <div key={student.idSiswa} className="space-y-8">
        {printScope === 'batch' && (
          <div className="no-print flex items-center justify-between px-3 py-1.5 bg-slate-800 text-white rounded-lg text-xs font-bold">
            <span>
              Buku Induk #{studentIdx + 1}: {student.namaSiswa} (NIS: {nisDisplay} · NISN: {nisnDisplay})
            </span>
            <span>Kelas {student.kelas}</span>
          </div>
        )}

        {/* ============================================================ */}
        {/* HALAMAN 1: KAVER BUKU INDUK PESERTA DIDIK                    */}
        {/* ============================================================ */}
        {showKaver && (
          <div
            className={`report-paper-sheet report-typography-scope ${sheetMaxWClass} mx-auto bg-white text-black shadow-lg border border-slate-300 p-6 sm:p-10 print:shadow-none print:border-none print:p-0 printable-area print-page-break`}
            style={{
              fontFamily: reportFontCss,
              fontSize: printFontSize,
              lineHeight: printLineHeight
            }}
          >
            <div className="w-full border-[3px] border-black p-1 box-border flex flex-col justify-between min-h-[255mm]">
              <div className="w-full flex-1 border border-black relative p-6 sm:p-8 flex flex-col justify-between bg-white">
                {/* 4 Sudut Ornamen Bingkai Klasik */}
                <svg className="absolute -top-1 -left-1 w-11 h-11 text-black pointer-events-none" viewBox="0 0 100 100" fill="currentColor">
                  <path d="M0,0 L100,0 C70,10 60,30 50,50 C40,70 20,80 0,100 Z" opacity="0.9" />
                  <circle cx="18" cy="18" r="6" />
                  <circle cx="36" cy="12" r="4" />
                  <circle cx="12" cy="36" r="4" />
                </svg>
                <svg className="absolute -top-1 -right-1 w-11 h-11 text-black pointer-events-none" viewBox="0 0 100 100" fill="currentColor">
                  <path d="M100,0 L0,0 C30,10 40,30 50,50 C60,70 80,80 100,100 Z" opacity="0.9" />
                  <circle cx="82" cy="18" r="6" />
                  <circle cx="64" cy="12" r="4" />
                  <circle cx="88" cy="36" r="4" />
                </svg>
                <svg className="absolute -bottom-1 -left-1 w-11 h-11 text-black pointer-events-none" viewBox="0 0 100 100" fill="currentColor">
                  <path d="M0,100 L100,100 C70,90 60,70 50,50 C40,30 20,20 0,0 Z" opacity="0.9" />
                  <circle cx="18" cy="82" r="6" />
                  <circle cx="36" cy="88" r="4" />
                  <circle cx="12" cy="64" r="4" />
                </svg>
                <svg className="absolute -bottom-1 -right-1 w-11 h-11 text-black pointer-events-none" viewBox="0 0 100 100" fill="currentColor">
                  <path d="M100,100 L0,100 C30,90 40,70 50,50 C60,30 80,20 100,0 Z" opacity="0.9" />
                  <circle cx="82" cy="82" r="6" />
                  <circle cx="64" cy="88" r="4" />
                  <circle cx="88" cy="64" r="4" />
                </svg>

                {/* Kotak Nomor Induk Atas */}
                <div className="flex justify-center pt-1">
                  <div className="border-2 border-black px-5 py-1.5 text-center bg-white">
                    <div className="text-[11px] font-black uppercase tracking-wider">
                      KURIKULUM MERDEKA · DOKUMEN ARSIP SEKOLAH
                    </div>
                    <div className="text-xs font-bold font-mono mt-0.5">
                      NO. INDUK (NIS): {nisDisplay} &nbsp;|&nbsp; NISN: {nisnDisplay}
                    </div>
                  </div>
                </div>

                {/* Logo & Judul Utama */}
                <div className="text-center my-4 space-y-3">
                  <div className="flex justify-center">
                    <LogoTutWuriHandayani
                      url={db.sekolah.logoKananUrl}
                      size={92}
                      className="w-22 h-22 object-contain"
                      label="Logo Tut Wuri Handayani"
                    />
                  </div>
                  <div className="space-y-1">
                    <h1 className="text-2xl sm:text-3xl font-black tracking-widest uppercase text-black">
                      BUKU INDUK PESERTA DIDIK
                    </h1>
                    <h2 className="text-lg sm:text-xl font-black tracking-wider uppercase text-black">
                      SEKOLAH DASAR (SD)
                    </h2>
                    <p className="text-sm font-bold uppercase tracking-wide text-black">
                      {db.sekolah.namaSekolah || 'SD NEGERI BALIL'}
                    </p>
                  </div>
                </div>

                {/* Biodata Ringkas Kaver (8 Baris Rapi Sejajar Titik Dua) */}
                <div className="max-w-lg w-full mx-auto space-y-2 text-xs sm:text-sm my-4 px-2">
                  {[
                    ['Nama Sekolah', db.sekolah.namaSekolah || '-'],
                    ['NPSN / NSS', `${db.sekolah.npsn || '-'} / ${db.sekolah.nss || '-'}`],
                    ['Nama Peserta Didik', student.namaSiswa.toUpperCase()],
                    ['NISN / NIS', `${nisnDisplay} / ${nisDisplay}`],
                    ['Kelas / Fase', `Kelas ${student.kelas} (Fase ${student.fase || getFaseFromKelas(student.kelas)})`],
                    ['Tahun Pelajaran', student.tahunPelajaran || db.sekolah.tahunPelajaran || '-'],
                    ['Alamat Sekolah', cleanVal(db.sekolah.alamat)],
                    [
                      'Kabupaten / Provinsi',
                      `${cleanVal(db.sekolah.kabupaten)}, ${cleanVal(db.sekolah.provinsi)}`
                    ]
                  ].map(([lbl, val], i) => (
                    <div key={i} className="flex items-start">
                      <span className="w-[12.5em] font-bold shrink-0">{lbl}</span>
                      <span className="w-[1.2em] text-center font-bold shrink-0">:</span>
                      <span className={`flex-1 ${i === 2 ? 'font-black uppercase' : 'font-semibold'}`}>
                        {val}
                      </span>
                    </div>
                  ))}
                </div>

                {/* Kementerian Bawah */}
                <div className="text-center pb-1 pt-4">
                  <div className="text-xs sm:text-sm font-black uppercase tracking-wide text-black">
                    KEMENTERIAN PENDIDIKAN, KEBUDAYAAN, RISET, DAN TEKNOLOGI
                  </div>
                  <div className="text-xs sm:text-sm font-bold uppercase tracking-widest text-black mt-0.5">
                    REPUBLIK INDONESIA
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ============================================================ */}
        {/* HALAMAN 2: PETUNJUK PENGGUNAAN & KEADAAN KELAS SEKOLAH       */}
        {/* ============================================================ */}
        {showPetunjuk && (
          <div
            className={`report-paper-sheet report-typography-scope ${sheetMaxWClass} mx-auto bg-white text-black shadow-lg border border-slate-300 p-6 sm:p-10 print:shadow-none print:border-none print:p-0 printable-area print-page-break`}
            style={{
              fontFamily: reportFontCss,
              fontSize: printFontSize,
              lineHeight: printLineHeight
            }}
          >
            <div className="space-y-4 text-xs leading-relaxed">
              {/* Header Bar */}
              <div className="border-y-2 border-black py-1.5 text-center bg-[#f5f5f5]">
                <h2 className="text-sm sm:text-base font-black uppercase tracking-wider">
                  PETUNJUK PENGISIAN DAN PENGGUNAAN BUKU INDUK PESERTA DIDIK
                </h2>
              </div>

              {/* A. Petunjuk Umum */}
              <div>
                <h3 className="font-black uppercase mb-1">A. Petunjuk Umum</h3>
                <ol className="list-decimal pl-5 space-y-1">
                  <li>
                    Buku Induk Peserta Didik merupakan dokumen resmi administrasi sekolah yang mencatat identitas lengkap, riwayat kesehatan, pendidikan, orang tua/wali, serta perkembangan hasil belajar setiap peserta didik selama menempuh pendidikan di satuan pendidikan.
                  </li>
                  <li>
                    Pengisian data peserta didik wajib disesuaikan dengan dokumen resmi yang sah (Akta Kelahiran, Kartu Keluarga, Ijazah/STTB sebelumnya, dan data pokok pendidikan/Dapodik).
                  </li>
                  <li>
                    Setiap peserta didik diberikan Nomor Induk Siswa (NIS) yang bersifat permanen selama bersekolah dan dilengkapi dengan Nomor Induk Siswa Nasional (NISN).
                  </li>
                  <li>
                    Lembar catatan hasil belajar diperbarui secara berkala setiap akhir semester dan akhir tahun pelajaran berdasarkan rapor resmi yang telah disahkan oleh Wali Kelas dan Kepala Sekolah.
                  </li>
                </ol>
              </div>

              {/* B. Ketentuan Pengisian */}
              <div>
                <h3 className="font-black uppercase mb-1">B. Ketentuan Pengisian &amp; Pengarsipan</h3>
                <ul className="list-disc pl-5 space-y-1">
                  <li>Pas foto ukuran 3 × 4 cm dipasang pada lembar identitas induk peserta didik (latar belakang merah/biru sesuai ketentuan sekolah).</li>
                  <li>Bagian A sampai H diisi lengkap meliputi keterangan pribadi, tempat tinggal, kesehatan, riwayat pendidikan, data ayah, ibu, wali, kegemaran, serta keterangan tamat belajar atau mutasi pindah sekolah.</li>
                  <li>Data nilai mata pelajaran wajib, muatan lokal, ekstrakurikuler, ketidakhadiran, dan keputusan kenaikan kelas disinkronkan langsung dari basis data penilaian Kurikulum Merdeka.</li>
                  <li>Buku Induk wajib disimpan sebagai arsip permanen satuan pendidikan dan dapat dicetak ulang sewaktu-waktu apabila dibutuhkan untuk verifikasi ijazah atau surat keterangan pindah.</li>
                </ul>
              </div>

              {/* C. Data Keadaan Rombongan Belajar Sekolah (2 Kolom Kelas I-III & IV-VI) */}
              <div className="pt-2 border-t border-black">
                <h3 className="font-black uppercase mb-2.5">
                  C. Data Keadaan Rombongan Belajar &amp; Wali Kelas ({db.sekolah.namaSekolah})
                </h3>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  {SD_CLASSES.map((kls) => {
                    const waliKls = getWaliForKelas(kls);
                    const studentsInKls = db.siswa.filter((s) => s.kelas === kls);
                    const countL = studentsInKls.filter((s) => s.jenisKelamin === 'L').length;
                    const countP = studentsInKls.filter((s) => s.jenisKelamin === 'P').length;
                    return (
                      <div key={kls} className="border border-black p-2.5 bg-white">
                        <div className="font-black uppercase border-b border-black pb-1 mb-1.5 flex items-center justify-between">
                          <span>Kelas {kls} (Fase {getFaseFromKelas(kls)})</span>
                          <span className="font-mono text-[11px]">{db.sekolah.tahunPelajaran}</span>
                        </div>
                        <div className="space-y-0.5 text-[11px]">
                          <div className="flex">
                            <span className="w-[7.5em] shrink-0">Wali Kelas</span>
                            <span className="w-[1em] text-center shrink-0">:</span>
                            <span className="font-bold flex-1">{waliKls.namaGuru}</span>
                          </div>
                          <div className="flex">
                            <span className="w-[7.5em] shrink-0">NIP</span>
                            <span className="w-[1em] text-center shrink-0">:</span>
                            <span className="font-mono flex-1">{waliKls.nip}</span>
                          </div>
                          <div className="flex">
                            <span className="w-[7.5em] shrink-0">Jumlah Siswa</span>
                            <span className="w-[1em] text-center shrink-0">:</span>
                            <span className="font-semibold flex-1">
                              {studentsInKls.length} Orang (L: {countL}, P: {countP})
                            </span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Tanda Tangan Kepala Sekolah */}
              <div className="pt-6 flex justify-end">
                <div className="w-64 text-left text-xs leading-snug">
                  <div>{placeAndDate}</div>
                  <div className="font-bold mt-0.5">Kepala Sekolah,</div>
                  <div className="h-16" />
                  <div className="font-black underline uppercase">
                    {db.sekolah.namaKepalaSekolah || '........................................'}
                  </div>
                  <div className="font-mono text-[11px]">
                    NIP. {cleanVal(db.sekolah.nipKepalaSekolah)}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ============================================================ */}
        {/* HALAMAN 3: LEMBAR IDENTITAS PESERTA DIDIK LENGKAP (A - H)    */}
        {/* ============================================================ */}
        {showBiodata && (
          <div
            className={`report-paper-sheet report-typography-scope ${sheetMaxWClass} mx-auto bg-white text-black shadow-lg border border-slate-300 p-6 sm:p-10 print:shadow-none print:border-none print:p-0 printable-area print-page-break`}
            style={{
              fontFamily: reportFontCss,
              fontSize: printFontSize,
              lineHeight: printLineHeight
            }}
          >
            {/* Header Judul */}
            <div className="text-center border-b-2 border-black pb-2 mb-3">
              <h2 className="text-base sm:text-lg font-black uppercase tracking-wider">
                LEMBAR BUKU INDUK PESERTA DIDIK
              </h2>
              <div className="text-xs font-bold font-mono mt-0.5">
                NOMOR INDUK SISWA (NIS) : {nisDisplay} &nbsp;&nbsp;|&nbsp;&nbsp; NISN : {nisnDisplay}
              </div>
            </div>

            {/* Atas: Bagian A + Pas Foto 3x4 di kanan */}
            <div className="flex flex-col sm:flex-row gap-4 items-start justify-between border-b border-dashed border-black pb-3 mb-3">
              <div className="flex-1 w-full text-xs space-y-1">
                <h3 className="font-black uppercase text-black mb-1">
                  A. KETERANGAN PRIBADI PESERTA DIDIK
                </h3>
                {[
                  ['1. Nama Lengkap', student.namaSiswa.toUpperCase()],
                  ['2. Nama Panggilan', cleanVal(bi.namaPanggilan)],
                  ['3. Jenis Kelamin', student.jenisKelamin === 'L' ? 'Laki-laki' : 'Perempuan'],
                  [
                    '4. Tempat, Tanggal Lahir',
                    `${cleanVal(student.tempatLahir)}, ${formatDateIndo(student.tanggalLahir)}`
                  ],
                  ['5. Agama', cleanVal(student.agama)],
                  ['6. Kewarganegaraan', cleanVal(bi.kewarganegaraan, 'Indonesia')],
                  ['7. Anak ke-', cleanVal(bi.anakKe)],
                  [
                    '8. Jumlah Saudara',
                    `Kandung: ${cleanVal(bi.jumlahSaudaraKandung, '0')}, Tiri: ${cleanVal(bi.jumlahSaudaraTiri, '0')}, Angkat: ${cleanVal(bi.jumlahSaudaraAngkat, '0')}`
                  ],
                  ['9. Yatim / Piatu / Lengkap', cleanVal(bi.anakYatim, 'Lengkap')],
                  ['10. Bahasa Sehari-hari', cleanVal(bi.bahasaSehariHari)]
                ].map(([label, val], idx) => (
                  <div key={idx} className="flex">
                    <span className="w-[15.5em] shrink-0">{label}</span>
                    <span className="w-[1.1em] text-center shrink-0">:</span>
                    <span className={`flex-1 ${idx === 0 ? 'font-bold uppercase' : ''}`}>{val}</span>
                  </div>
                ))}
              </div>

              {/* Kotak Pas Foto 3 x 4 */}
              <div className="shrink-0 flex flex-col items-center sm:pt-4 mx-auto sm:mx-0">
                <div className="w-[28mm] h-[38mm] border-2 border-black flex items-center justify-center bg-slate-50 overflow-hidden">
                  {student.fotoUrl ? (
                    <img
                      src={student.fotoUrl}
                      alt={`Pas Foto ${student.namaSiswa}`}
                      className="w-full h-full object-cover"
                    />
                  ) : (
                    <span className="text-[10px] text-center text-slate-500 font-sans px-2">
                      Pas Foto
                      <br />3 × 4 cm
                    </span>
                  )}
                </div>
              </div>
            </div>

            {/* B. Keterangan Tempat Tinggal */}
            <div className="text-xs space-y-1 border-b border-dashed border-black pb-2.5 mb-2.5">
              <h3 className="font-black uppercase text-black mb-1">B. KETERANGAN TEMPAT TINGGAL</h3>
              {[
                ['11. Alamat Lengkap', cleanVal(student.alamat)],
                ['12. Nomor Telepon / HP', cleanVal(student.telepon)],
                ['13. Tinggal Dengan', cleanVal(bi.tinggalDengan, 'Orang Tua')],
                ['14. Jarak / Transportasi ke Sekolah', [bi.jarakKeSekolah, bi.transportasi].filter(Boolean).join(' · ') || '-']
              ].map(([label, val], idx) => (
                <div key={idx} className="flex">
                  <span className="w-[17em] shrink-0">{label}</span>
                  <span className="w-[1.1em] text-center shrink-0">:</span>
                  <span className="flex-1">{val}</span>
                </div>
              ))}
            </div>

            {/* C. Keterangan Kesehatan */}
            <div className="text-xs space-y-1 border-b border-dashed border-black pb-2.5 mb-2.5">
              <h3 className="font-black uppercase text-black mb-1">C. KETERANGAN KESEHATAN</h3>
              {[
                ['15. Golongan Darah', cleanVal(bi.golonganDarah)],
                ['16. Penyakit yang Pernah Diderita', cleanVal(bi.penyakitPernahDiderita)],
                ['17. Kelainan Jasmani', cleanVal(bi.kelainanJasmani)],
                [
                  '18. Tinggi / Berat Badan',
                  `Tinggi: ${cleanVal(bi.tinggiBadan, '___')} cm   |   Berat: ${cleanVal(bi.beratBadan, '___')} kg`
                ]
              ].map(([label, val], idx) => (
                <div key={idx} className="flex">
                  <span className="w-[17em] shrink-0">{label}</span>
                  <span className="w-[1.1em] text-center shrink-0">:</span>
                  <span className="flex-1">{val}</span>
                </div>
              ))}
            </div>

            {/* D. Keterangan Pendidikan */}
            <div className="text-xs space-y-1 border-b border-dashed border-black pb-2.5 mb-2.5">
              <h3 className="font-black uppercase text-black mb-1">D. KETERANGAN PENDIDIKAN</h3>
              {[
                ['19. Pendidikan Sebelumnya (TK/PAUD/SD)', cleanVal(bi.asalTkPaud || student.sekolahAsal)],
                ['    Tanggal & No. STTB / Ijazah', cleanVal(bi.tanggalNoSttbSebelumnya)],
                ['20. Pindahan dari Sekolah / Alasan', [bi.pindahanDariSekolah, bi.alasanPindah].filter(Boolean).join(' — ') || '-'],
                [
                  '21. Diterima di Sekolah Ini',
                  `Kelas ${cleanVal(student.diterimaKelas, student.kelas)}   ·   Tanggal: ${formatDateIndo(student.diterimaTanggal)}`
                ]
              ].map(([label, val], idx) => (
                <div key={idx} className="flex">
                  <span className="w-[17em] shrink-0 whitespace-pre">{label}</span>
                  <span className="w-[1.1em] text-center shrink-0">:</span>
                  <span className="flex-1">{val}</span>
                </div>
              ))}
            </div>

            {/* E & F: Ayah & Ibu Kandung (2 Kolom Sejajar Rapi agar Muat 1 Halaman) */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs border-b border-dashed border-black pb-2.5 mb-2.5">
              {/* E. Ayah Kandung */}
              <div className="space-y-1">
                <h3 className="font-black uppercase text-black mb-1">E. KETERANGAN AYAH KANDUNG</h3>
                {[
                  ['22. Nama Ayah', cleanVal(student.namaAyah)],
                  ['    Tempat, Tgl Lahir', cleanVal(bi.ttlAyah)],
                  ['    Agama / Kewarganegaraan', `${cleanVal(bi.agamaAyah)} / ${cleanVal(bi.kewarganegaraanAyah, 'Indonesia')}`],
                  ['    Pendidikan Terakhir', cleanVal(bi.pendidikanAyah)],
                  ['    Pekerjaan', cleanVal(student.pekerjaanAyah)],
                  ['    Penghasilan / Bulan', cleanVal(bi.penghasilanAyah)],
                  ['    Alamat / No. Telp', [bi.alamatAyah || student.alamat, bi.teleponAyah].filter((v) => v && v !== '-').join(' · ') || '-'],
                  ['    Status', cleanVal(bi.statusAyah, 'Masih Hidup')]
                ].map(([label, val], idx) => (
                  <div key={idx} className="flex">
                    <span className="w-[12.5em] shrink-0 whitespace-pre">{label}</span>
                    <span className="w-[1em] text-center shrink-0">:</span>
                    <span className={`flex-1 ${idx === 0 ? 'font-bold uppercase' : ''}`}>{val}</span>
                  </div>
                ))}
              </div>

              {/* F. Ibu Kandung */}
              <div className="space-y-1">
                <h3 className="font-black uppercase text-black mb-1">F. KETERANGAN IBU KANDUNG</h3>
                {[
                  ['23. Nama Ibu', cleanVal(student.namaIbu)],
                  ['    Tempat, Tgl Lahir', cleanVal(bi.ttlIbu)],
                  ['    Agama / Kewarganegaraan', `${cleanVal(bi.agamaIbu)} / ${cleanVal(bi.kewarganegaraanIbu, 'Indonesia')}`],
                  ['    Pendidikan Terakhir', cleanVal(bi.pendidikanIbu)],
                  ['    Pekerjaan', cleanVal(student.pekerjaanIbu)],
                  ['    Penghasilan / Bulan', cleanVal(bi.penghasilanIbu)],
                  ['    Alamat / No. Telp', [bi.alamatIbu || student.alamat, bi.teleponIbu].filter((v) => v && v !== '-').join(' · ') || '-'],
                  ['    Status', cleanVal(bi.statusIbu, 'Masih Hidup')]
                ].map(([label, val], idx) => (
                  <div key={idx} className="flex">
                    <span className="w-[12.5em] shrink-0 whitespace-pre">{label}</span>
                    <span className="w-[1em] text-center shrink-0">:</span>
                    <span className={`flex-1 ${idx === 0 ? 'font-bold uppercase' : ''}`}>{val}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* G & H: Wali & Perkembangan / Meninggalkan Sekolah */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs pb-2">
              <div className="space-y-1">
                <h3 className="font-black uppercase text-black mb-1">G. KETERANGAN WALI</h3>
                {[
                  ['24. Nama Wali', cleanVal(student.namaWali)],
                  ['    Hubungan Keluarga', cleanVal(bi.hubunganWali)],
                  ['    Pendidikan Terakhir', cleanVal(bi.pendidikanWali)],
                  ['    Pekerjaan', cleanVal(student.pekerjaanWali)],
                  ['    Penghasilan / Bulan', cleanVal(bi.penghasilanWali)],
                  ['    Alamat / No. Telp', [student.alamatWali, bi.teleponWali].filter((v) => v && v !== '-').join(' · ') || '-']
                ].map(([label, val], idx) => (
                  <div key={idx} className="flex">
                    <span className="w-[12.5em] shrink-0 whitespace-pre">{label}</span>
                    <span className="w-[1em] text-center shrink-0">:</span>
                    <span className="flex-1">{val}</span>
                  </div>
                ))}
              </div>

              <div className="space-y-1">
                <h3 className="font-black uppercase text-black mb-1">
                  H. KEGEMARAN &amp; MENINGGALKAN SEKOLAH
                </h3>
                {[
                  [
                    '25. Kegemaran / Bakat',
                    [
                      bi.hobiKesenian ? `Seni: ${bi.hobiKesenian}` : '',
                      bi.hobiOlahraga ? `Olahraga: ${bi.hobiOlahraga}` : '',
                      bi.hobiOrganisasi ? `Organisasi: ${bi.hobiOrganisasi}` : '',
                      bi.hobiLain
                    ]
                      .filter(Boolean)
                      .join('; ') || '-'
                  ],
                  ['26. Beasiswa / Prestasi', cleanVal(bi.beasiswa)],
                  ['27. Tamat Belajar / No. Ijazah', [formatDateIndo(bi.tamatBelajarTanggal), bi.noIjazahSttb].filter((v) => v && v !== '-').join(' · ') || '-'],
                  ['    Pindah Sekolah / Alasan', [formatDateIndo(bi.tanggalMeninggalkan), bi.alasanMeninggalkan].filter((v) => v && v !== '-').join(' · ') || '-'],
                  ['    Melanjutkan ke', cleanVal(bi.melanjutkanKe)],
                  ['    Catatan Khusus', cleanVal(bi.catatanKhusus)]
                ].map(([label, val], idx) => (
                  <div key={idx} className="flex">
                    <span className="w-[12.5em] shrink-0 whitespace-pre">{label}</span>
                    <span className="w-[1em] text-center shrink-0">:</span>
                    <span className="flex-1">{val}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Pengesahan Kepala Sekolah */}
            <div className="pt-4 mt-2 border-t border-black flex justify-end">
              <div className="w-64 text-left text-xs leading-snug">
                <div>{placeAndDate}</div>
                <div className="font-bold mt-0.5">Kepala Sekolah,</div>
                <div className="h-14" />
                <div className="font-black underline uppercase">
                  {db.sekolah.namaKepalaSekolah || '........................................'}
                </div>
                <div className="font-mono text-[11px]">
                  NIP. {cleanVal(db.sekolah.nipKepalaSekolah)}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ============================================================ */}
        {/* HALAMAN 4: TRANSKRIP CATATAN HASIL BELAJAR BUKU INDUK         */}
        {/* ============================================================ */}
        {showNilai && (
          <div
            className={`report-paper-sheet report-typography-scope ${sheetMaxWClass} mx-auto bg-white text-black shadow-lg border border-slate-300 p-6 sm:p-10 print:shadow-none print:border-none print:p-0 printable-area print-page-break`}
            style={{
              fontFamily: reportFontCss,
              fontSize: printFontSize,
              lineHeight: printLineHeight
            }}
          >
            {/* Header Judul */}
            <div className="text-center mb-4">
              <h2 className="text-base font-black uppercase tracking-wider">
                CATATAN HASIL BELAJAR BUKU INDUK PESERTA DIDIK
              </h2>
              <p className="text-xs font-semibold">
                {db.sekolah.namaSekolah} · Semester {!isSem2 ? '1 (Ganjil)' : '2 (Genap)'} · Tahun Pelajaran {db.sekolah.tahunPelajaran}
              </p>
            </div>

            {/* Identitas 2 Kolom */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-8 gap-y-1 text-xs mb-3">
              <div className="space-y-1">
                <div className="flex">
                  <span className="w-[10.5em] shrink-0">Nama Peserta Didik</span>
                  <span className="w-[1em] text-center shrink-0">:</span>
                  <span className="font-bold uppercase flex-1">{student.namaSiswa}</span>
                </div>
                <div className="flex">
                  <span className="w-[10.5em] shrink-0">NISN / NIS</span>
                  <span className="w-[1em] text-center shrink-0">:</span>
                  <span className="font-mono flex-1">{nisnDisplay} / {nisDisplay}</span>
                </div>
                <div className="flex">
                  <span className="w-[10.5em] shrink-0">Sekolah</span>
                  <span className="w-[1em] text-center shrink-0">:</span>
                  <span className="flex-1">{db.sekolah.namaSekolah}</span>
                </div>
              </div>
              <div className="space-y-1">
                <div className="flex">
                  <span className="w-[9.5em] shrink-0">Kelas / Fase</span>
                  <span className="w-[1em] text-center shrink-0">:</span>
                  <span className="flex-1">Kelas {student.kelas} / Fase {student.fase || getFaseFromKelas(student.kelas)}</span>
                </div>
                <div className="flex">
                  <span className="w-[9.5em] shrink-0">Semester</span>
                  <span className="w-[1em] text-center shrink-0">:</span>
                  <span className="flex-1">{!isSem2 ? '1 (Ganjil)' : '2 (Genap)'}</span>
                </div>
                <div className="flex">
                  <span className="w-[9.5em] shrink-0">Tahun Pelajaran</span>
                  <span className="w-[1em] text-center shrink-0">:</span>
                  <span className="flex-1">{db.sekolah.tahunPelajaran}</span>
                </div>
              </div>
            </div>

            {/* Tabel Utama Nilai Akademik */}
            <table className="w-full text-[11px] border-collapse border border-black mb-3">
              <thead>
                <tr className="bg-[#f2f2f2] font-bold">
                  <th className="border border-black py-1.5 px-1.5 text-center w-[2.8em]">No</th>
                  <th className="border border-black py-1.5 px-2 text-left w-[26%]">Mata Pelajaran</th>
                  <th className="border border-black py-1.5 px-1.5 text-center w-[4em]">Nilai</th>
                  <th className="border border-black py-1.5 px-2 text-left">Capaian Kompetensi / Deskripsi</th>
                </tr>
              </thead>
              <tbody>
                {wajibScores.map((item, idx) => (
                  <tr key={item.mapel.kode || item.mapel.mataPelajaran}>
                    <td className="border border-black py-1.5 px-1.5 text-center align-middle">{idx + 1}</td>
                    <td className="border border-black py-1.5 px-2 font-bold align-middle">{item.namaTampil}</td>
                    <td className="border border-black py-1.5 px-1.5 text-center font-bold align-middle bg-[#fcf8e8]">
                      {item.nilaiAkhir > 0 ? item.nilaiAkhir : '-'}
                    </td>
                    <td
                      className="border border-black py-1.5 px-2 align-middle leading-snug"
                      style={{ textAlign: printTextAlign }}
                    >
                      {item.deskripsi}
                    </td>
                  </tr>
                ))}

                {pilihanMulokScores.length > 0 && (
                  <>
                    <tr className="bg-slate-50 font-bold">
                      <td className="border border-black py-1 px-1.5 text-center">#</td>
                      <td colSpan={3} className="border border-black py-1 px-2 uppercase text-[10px]">
                        Muatan Lokal &amp; Mata Pelajaran Pilihan
                      </td>
                    </tr>
                    {pilihanMulokScores.map((item, idx) => (
                      <tr key={item.mapel.kode || item.mapel.mataPelajaran}>
                        <td className="border border-black py-1.5 px-1.5 text-center align-middle">
                          {String.fromCharCode(97 + idx)}
                        </td>
                        <td className="border border-black py-1.5 px-2 font-bold align-middle pl-4">
                          {item.namaTampil}
                        </td>
                        <td className="border border-black py-1.5 px-1.5 text-center font-bold align-middle bg-[#fcf8e8]">
                          {item.nilaiAkhir > 0 ? item.nilaiAkhir : '-'}
                        </td>
                        <td
                          className="border border-black py-1.5 px-2 align-middle leading-snug"
                          style={{ textAlign: printTextAlign }}
                        >
                          {item.deskripsi}
                        </td>
                      </tr>
                    ))}
                  </>
                )}

                <tr className="bg-[#f9f9f9] font-bold">
                  <td colSpan={2} className="border border-black py-1.5 px-2 text-right">
                    Jumlah Nilai / Rata-Rata
                  </td>
                  <td className="border border-black py-1.5 px-1.5 text-center font-black">
                    {scoredItems.length > 0 ? totalNilai : '-'}
                  </td>
                  <td className="border border-black py-1.5 px-2 text-left">
                    Rata-Rata Nilai: <span className="font-black">{rataRata}</span> ({scoredItems.length} Mata Pelajaran Dinilai)
                  </td>
                </tr>
              </tbody>
            </table>

            {/* Tabel Ekstrakurikuler */}
            <table className="w-full text-[11px] border-collapse border border-black mb-3">
              <thead>
                <tr className="bg-[#f2f2f2] font-bold">
                  <th className="border border-black py-1 px-1.5 text-center w-[2.8em]">No</th>
                  <th className="border border-black py-1 px-2 text-left w-[30%]">Ekstrakurikuler</th>
                  <th className="border border-black py-1 px-2 text-left">Predikat &amp; Keterangan</th>
                </tr>
              </thead>
              <tbody>
                {studentEkskul.length === 0 ? (
                  <tr>
                    <td className="border border-black py-1.5 px-1.5 text-center">1</td>
                    <td className="border border-black py-1.5 px-2">-</td>
                    <td className="border border-black py-1.5 px-2">-</td>
                  </tr>
                ) : (
                  studentEkskul.map((e, i) => (
                    <tr key={e.id || i}>
                      <td className="border border-black py-1 px-1.5 text-center">{i + 1}</td>
                      <td className="border border-black py-1 px-2 font-bold">{e.jenisEkstrakurikuler}</td>
                      <td className="border border-black py-1 px-2">
                        <span className="font-semibold">{e.predikat}</span>
                        {e.keterangan ? ` — ${e.keterangan}` : ''}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>

            {/* Ketidakhadiran & Catatan / Keputusan */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4 text-[11px]">
              <table className="w-full border-collapse border border-black">
                <thead>
                  <tr className="bg-[#f2f2f2]">
                    <th colSpan={3} className="border border-black py-1 px-2 text-left font-bold">
                      Ketidakhadiran
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td className="border border-black py-1 px-2 w-[9em]">Sakit</td>
                    <td className="border border-black py-1 px-1 text-center w-[1em]">:</td>
                    <td className="border border-black py-1 px-2">{studentAbsen.sakit || 0} hari</td>
                  </tr>
                  <tr>
                    <td className="border border-black py-1 px-2">Izin</td>
                    <td className="border border-black py-1 px-1 text-center">:</td>
                    <td className="border border-black py-1 px-2">{studentAbsen.izin || 0} hari</td>
                  </tr>
                  <tr>
                    <td className="border border-black py-1 px-2">Tanpa Keterangan</td>
                    <td className="border border-black py-1 px-1 text-center">:</td>
                    <td className="border border-black py-1 px-2">{studentAbsen.alpa || 0} hari</td>
                  </tr>
                </tbody>
              </table>

              <div className="border border-black p-2 space-y-1.5">
                <div className="font-bold border-b border-black pb-0.5">
                  Catatan Wali Kelas &amp; Keputusan
                </div>
                <p className="leading-snug min-h-[28px]">
                  {studentNote?.catatanWaliKelas || 'Terus tingkatkan semangat belajar dan kedisiplinan.'}
                </p>
                {isSem2 && (
                  <div className="pt-1 border-t border-dashed border-black font-bold">
                    Keputusan:{' '}
                    {studentDecision?.statusKeputusan === 'Lulus'
                      ? 'LULUS dari Satuan Pendidikan'
                      : studentDecision?.statusKeputusan === 'Tinggal di Kelas'
                      ? `Tinggal di Kelas ${studentDecision.tinggalKelas || student.kelas}`
                      : `Naik ke Kelas ${studentDecision?.naikKelas || 'Berikutnya'}`}
                  </div>
                )}
              </div>
            </div>

            {/* Tanda Tangan 3 Pejabat / Orang Tua */}
            <div className="grid grid-cols-3 gap-4 pt-3 text-xs leading-snug">
              <div className="text-left">
                <div>Mengetahui,</div>
                <div className="font-bold">Orang Tua / Wali</div>
                <div className="h-14" />
                <div className="font-bold underline uppercase">
                  {cleanVal(student.namaAyah || student.namaWali, '..........................')}
                </div>
              </div>
              <div className="text-left">
                <div>&nbsp;</div>
                <div className="font-bold">Wali Kelas {student.kelas},</div>
                <div className="h-14" />
                <div className="font-bold underline uppercase">{wali.namaGuru}</div>
                <div className="font-mono text-[10px]">NIP. {wali.nip}</div>
              </div>
              <div className="text-left">
                <div>{placeAndDate}</div>
                <div className="font-bold">Kepala Sekolah,</div>
                <div className="h-14" />
                <div className="font-bold underline uppercase">
                  {db.sekolah.namaKepalaSekolah || '........................................'}
                </div>
                <div className="font-mono text-[10px]">NIP. {cleanVal(db.sekolah.nipKepalaSekolah)}</div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="space-y-4">
      {/* Dynamic Print CSS */}
      <style>{`
        @media print {
          @page {
            size: ${finalPageCssSize};
            margin: ${paperMargins.top}mm ${paperMargins.right}mm ${paperMargins.bottom}mm ${paperMargins.left}mm;
          }
          html, body, #root, #root > div, main {
            overflow: visible !important;
            height: auto !important;
            max-height: none !important;
            position: static !important;
            width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
            background: #ffffff !important;
          }
          .print-page-break {
            page-break-after: always !important;
            break-after: page !important;
            margin-bottom: 0 !important;
            border: none !important;
            box-shadow: none !important;
          }
          .no-print, header, nav, aside, button:not(.allow-print) {
            display: none !important;
          }
          .printable-area {
            zoom: ${printScale / 100};
          }
        }
      `}</style>

      {/* 1. HEADER & BAR AKSI IKON RINGKAS */}
      <div className="bg-white dark:bg-slate-900 rounded-2xl p-3.5 sm:p-4 border border-slate-200 dark:border-slate-700 shadow-xs no-print space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2.5">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0">
              <BookUser size={19} />
            </div>
            <div>
              <h2 className="text-sm sm:text-base font-black text-slate-900 dark:text-white leading-tight">
                BUKU INDUK PESERTA DIDIK
              </h2>
              <p className="text-[11px] text-slate-500 dark:text-slate-400">
                Arsip permanen identitas lengkap (A–H), kaver resmi, data rombel, dan transkrip nilai siswa
              </p>
            </div>
          </div>

          {/* Tombol Aksi Ikon Saja (Ringkas Tanpa Memakan Tempat) */}
          <div className="flex flex-wrap items-center gap-1.5">
            {/* Ukuran Kertas Cepat */}
            <div className="inline-flex items-center bg-slate-100 dark:bg-slate-800 p-0.5 rounded-lg border border-slate-200 dark:border-slate-700 gap-0.5">
              {(['F4', 'A4', 'Legal'] as const).map((sz) => (
                <button
                  key={sz}
                  type="button"
                  onClick={() => {
                    setPaperSize(sz);
                    showSaveNotification(`Ukuran kertas diubah ke ${sz}`);
                  }}
                  title={`Kertas ${sz}`}
                  className={`px-2 h-7 rounded-md text-[11px] font-bold transition-all cursor-pointer ${
                    paperSize === sz
                      ? 'bg-amber-500 text-white shadow-2xs'
                      : 'text-slate-600 dark:text-slate-300 hover:bg-white'
                  }`}
                >
                  {sz}
                </button>
              ))}
              <button
                type="button"
                onClick={() => setPaperOrientation(paperOrientation === 'portrait' ? 'landscape' : 'portrait')}
                title={paperOrientation === 'portrait' ? 'Orientasi Vertikal (Klik untuk Horizontal)' : 'Orientasi Horizontal (Klik untuk Vertikal)'}
                aria-label="Ubah orientasi kertas"
                className="w-7 h-7 rounded-md inline-flex items-center justify-center bg-emerald-600 text-white cursor-pointer"
              >
                <span
                  className={`inline-block border-2 border-current rounded-[2px] ${
                    paperOrientation === 'portrait' ? 'w-2.5 h-3.5' : 'w-3.5 h-2.5'
                  }`}
                />
              </button>
            </div>

            {/* Huruf Cepat */}
            <div className="inline-flex items-center bg-blue-50 dark:bg-slate-800 p-0.5 rounded-lg border border-blue-200 dark:border-slate-700 gap-0.5">
              <button
                type="button"
                onClick={() => adjustFontSize('down')}
                title="Perkecil huruf"
                aria-label="Perkecil huruf"
                className="w-7 h-7 inline-flex items-center justify-center rounded-md bg-white hover:bg-slate-100 text-slate-800 text-xs font-black border border-slate-200 cursor-pointer"
              >
                A-
              </button>
              <span className="px-1.5 text-[11px] font-black text-blue-900 dark:text-blue-200">
                {printFontSize}
              </span>
              <button
                type="button"
                onClick={() => adjustFontSize('up')}
                title="Perbesar huruf"
                aria-label="Perbesar huruf"
                className="w-7 h-7 inline-flex items-center justify-center rounded-md bg-blue-600 hover:bg-blue-700 text-white text-xs font-black cursor-pointer"
              >
                A+
              </button>
            </div>

            {currentStudent && (
              <button
                type="button"
                onClick={() => handleOpenEditModal(currentStudent)}
                title={`Edit Biodata Lengkap Buku Induk (${currentStudent.namaSiswa})`}
                aria-label="Edit Biodata Buku Induk"
                className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-amber-500 hover:bg-amber-600 text-slate-950 shadow-xs cursor-pointer"
              >
                <Edit3 size={15} />
              </button>
            )}

            <button
              type="button"
              onClick={() => {
                const count = syncAllScoresWithTPDescriptions(transcriptSemester);
                showSaveNotification(`Berhasil menyinkronkan ${count} deskripsi capaian dari TP!`);
              }}
              title="Sinkronkan Deskripsi Nilai dari Tujuan Pembelajaran (TP)"
              aria-label="Sinkronkan TP"
              className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-blue-600 hover:bg-blue-700 text-white shadow-xs cursor-pointer"
            >
              <Zap size={15} />
            </button>

            <button
              type="button"
              onClick={() => setIsPreviewVisible((v) => !v)}
              title={isPreviewVisible ? 'Sembunyikan Lembar Kertas' : 'Tampilkan Lembar Kertas'}
              aria-label="Toggle Pratinjau"
              className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 cursor-pointer"
            >
              {isPreviewVisible ? <Eye size={15} /> : <EyeOff size={15} />}
            </button>

            <button
              type="button"
              onClick={handleExportExcel}
              title="Ekspor Buku Induk ke Excel (.xlsx)"
              aria-label="Ekspor Excel"
              className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white shadow-xs cursor-pointer"
            >
              <FileSpreadsheet size={15} />
            </button>

            <button
              type="button"
              onClick={handleExportHtmlPdf}
              title="Unduh Dokumen Siap Cetak / PDF (.html)"
              aria-label="Unduh PDF HTML"
              className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-slate-800 hover:bg-slate-900 text-white shadow-xs cursor-pointer"
            >
              <Download size={15} />
            </button>

            <button
              type="button"
              onClick={handlePrintBukuInduk}
              title="Cetak Buku Induk Siswa"
              aria-label="Cetak Buku Induk"
              className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-gradient-to-r from-emerald-600 to-teal-700 hover:from-emerald-500 hover:to-teal-600 text-white shadow-xs cursor-pointer"
            >
              <Printer size={15} className="text-amber-300" />
            </button>

            <button
              type="button"
              onClick={() => setFilterPanelMode((m) => (m === 'compact' ? 'full' : 'compact'))}
              title={filterPanelMode === 'compact' ? 'Perluas Filter' : 'Rampingkan Filter'}
              aria-label="Ubah mode filter"
              className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 cursor-pointer"
            >
              {filterPanelMode === 'compact' ? <Maximize2 size={14} /> : <Minimize2 size={14} />}
            </button>
          </div>
        </div>

        {/* Pengaturan Huruf & Kertas Lengkap */}
        <PrintTypographyControl />

        {/* Filter Siswa, Kelas, Semester & Tab Halaman Buku Induk */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-2.5 pt-1 text-xs">
          <div className="space-y-1">
            <label className="font-bold text-slate-700 dark:text-slate-300 block">Kelas &amp; Siswa:</label>
            <div className="flex items-center gap-1.5">
              <select
                value={selectedKelas}
                onChange={(e) => setSelectedKelas(e.target.value as ClassLevel)}
                className="px-2.5 py-1.5 border border-slate-300 rounded-xl bg-white font-bold text-slate-800 cursor-pointer"
              >
                {SD_CLASSES.map((k) => (
                  <option key={k} value={k}>
                    Kelas {k}
                  </option>
                ))}
              </select>
              <div className="flex-1">
                <StudentFilterSelect />
              </div>
            </div>
          </div>

          <div className="space-y-1">
            <label className="font-bold text-slate-700 dark:text-slate-300 block">
              Semester Transkrip &amp; Metode Cetak:
            </label>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={() => {
                  setTranscriptSemester('Ganjil');
                  setSelectedSemester('Ganjil');
                }}
                className={`flex-1 py-1.5 px-2 rounded-xl font-bold border cursor-pointer ${
                  transcriptSemester === 'Ganjil' || transcriptSemester === '1'
                    ? 'bg-blue-600 text-white border-blue-600'
                    : 'bg-slate-50 text-slate-700 border-slate-200'
                }`}
              >
                Sem 1 (Ganjil)
              </button>
              <button
                type="button"
                onClick={() => {
                  setTranscriptSemester('Genap');
                  setSelectedSemester('Genap');
                }}
                className={`flex-1 py-1.5 px-2 rounded-xl font-bold border cursor-pointer ${
                  transcriptSemester === 'Genap' || transcriptSemester === '2'
                    ? 'bg-emerald-600 text-white border-emerald-600'
                    : 'bg-slate-50 text-slate-700 border-slate-200'
                }`}
              >
                Sem 2 (Genap)
              </button>
              <button
                type="button"
                onClick={() => setPrintScope((s) => (s === 'single' ? 'batch' : 'single'))}
                title="Beralih antara Cetak 1 Siswa dan Cetak Kolektif Banyak Siswa"
                className={`px-2.5 py-1.5 rounded-xl font-bold border inline-flex items-center gap-1 cursor-pointer ${
                  printScope === 'batch'
                    ? 'bg-purple-700 text-white border-purple-700'
                    : 'bg-slate-100 text-slate-800 border-slate-300'
                }`}
              >
                {printScope === 'batch' ? <Users size={13} /> : <User size={13} />}
                <span>{printScope === 'batch' ? 'Kolektif' : '1 Siswa'}</span>
              </button>
            </div>
          </div>

          <div className="space-y-1">
            <label className="font-bold text-slate-700 dark:text-slate-300 block">Navigasi Siswa:</label>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                onClick={handlePrevStudent}
                title="Siswa Sebelumnya"
                className="p-1.5 border border-slate-300 rounded-xl bg-white hover:bg-slate-100 cursor-pointer"
              >
                <ChevronLeft size={15} />
              </button>
              <select
                value={currentStudent?.idSiswa || ''}
                onChange={(e) => setPrintTargetStudentId(e.target.value)}
                className="flex-1 px-2.5 py-1.5 border border-slate-300 rounded-xl bg-white font-bold text-slate-800 truncate cursor-pointer"
              >
                {filteredStudents.map((s) => (
                  <option key={s.idSiswa} value={s.idSiswa}>
                    {s.namaSiswa} ({cleanVal(s.nis)})
                  </option>
                ))}
              </select>
              <button
                type="button"
                onClick={handleNextStudent}
                title="Siswa Berikutnya"
                className="p-1.5 border border-slate-300 rounded-xl bg-white hover:bg-slate-100 cursor-pointer"
              >
                <ChevronRight size={15} />
              </button>
            </div>
          </div>
        </div>

        {/* Pilihan Halaman Buku Induk (1-4 / Semua) */}
        <div className="flex flex-wrap items-center justify-between gap-2 pt-1 border-t border-slate-100 dark:border-slate-800">
          <div className="flex flex-wrap items-center gap-1.5 text-xs">
            {[
              { id: 'semua', label: 'Semua (4 Halaman)', icon: <FileCheck2 size={13} /> },
              { id: 'kaver', label: '1. Kaver Buku Induk', icon: <BookOpen size={13} /> },
              { id: 'petunjuk', label: '2. Petunjuk & Rombel', icon: <FileText size={13} /> },
              { id: 'biodata', label: '3. Lembar Biodata (A–H)', icon: <User size={13} /> },
              { id: 'nilai', label: '4. Transkrip Nilai', icon: <GraduationCap size={13} /> }
            ].map((tab) => (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveTab(tab.id as BukuIndukPageTab)}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl font-bold border transition-all cursor-pointer ${
                  activeTab === tab.id
                    ? 'bg-slate-900 text-white border-slate-900 shadow-2xs'
                    : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`}
              >
                {tab.icon}
                <span>{tab.label}</span>
              </button>
            ))}
          </div>

          {filterPanelMode === 'full' && (
            <div className="relative min-w-[220px]">
              <Search size={13} className="absolute left-2.5 top-2 text-slate-400" />
              <input
                type="text"
                value={searchStudent}
                onChange={(e) => setSearchStudent(e.target.value)}
                placeholder="Cari nama / NIS / NISN..."
                className="w-full pl-7 pr-3 py-1.5 text-xs border border-slate-300 rounded-xl bg-slate-50"
              />
            </div>
          )}
        </div>

        {/* Batch Selector jika Kolektif aktif */}
        {printScope === 'batch' && (
          <div className="p-3 rounded-xl bg-purple-50 border border-purple-200 space-y-2 text-xs">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <span className="font-bold text-purple-950">
                Pilih Siswa untuk Cetak Kolektif ({selectedStudentIds.length}/{filteredStudents.length} terpilih):
              </span>
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => setSelectedStudentIds(filteredStudents.map((s) => s.idSiswa))}
                  className="px-2.5 py-1 bg-purple-700 text-white rounded-lg font-bold cursor-pointer"
                >
                  Pilih Semua
                </button>
                <button
                  type="button"
                  onClick={() => setSelectedStudentIds([])}
                  className="px-2.5 py-1 bg-white border border-purple-300 text-purple-900 rounded-lg font-bold cursor-pointer"
                >
                  Kosongkan
                </button>
              </div>
            </div>
            <div className="flex flex-wrap gap-1.5 max-h-28 overflow-y-auto">
              {filteredStudents.map((s) => {
                const checked = selectedStudentIds.includes(s.idSiswa);
                return (
                  <button
                    key={s.idSiswa}
                    type="button"
                    onClick={() => toggleBatchStudent(s.idSiswa)}
                    className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-lg border font-semibold cursor-pointer ${
                      checked
                        ? 'bg-purple-600 text-white border-purple-600'
                        : 'bg-white text-slate-700 border-slate-200'
                    }`}
                  >
                    {checked ? <CheckSquare size={12} /> : <Square size={12} />}
                    <span>{s.namaSiswa}</span>
                  </button>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* 2. AREA KERTAS PRATINJAU & CETAK BUKU INDUK */}
      {isPreviewVisible && (
        <div id="buku-induk-print-area" className="space-y-8 pb-10">
          {studentsToRender.length === 0 ? (
            <div className="bg-white rounded-2xl p-10 text-center border border-slate-200 text-slate-500 text-sm">
              Belum ada data siswa pada Kelas {selectedKelas}. Silakan tambahkan siswa di menu Data Siswa.
            </div>
          ) : (
            studentsToRender.map((student, idx) => renderStudentBukuInduk(student, idx))
          )}
        </div>
      )}

      {/* 3. MODAL EDIT BIODATA LENGKAP BUKU INDUK (A - H + PAS FOTO) */}
      {editModalOpen && editStudentTarget && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-3 sm:p-4 no-print">
          <div className="bg-white rounded-2xl shadow-2xl max-w-3xl w-full overflow-hidden border border-slate-200 flex flex-col max-h-[90vh]">
            <div className="bg-gradient-to-r from-emerald-700 to-teal-800 p-4 text-white flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2">
                <BookUser size={19} />
                <div>
                  <h3 className="font-bold text-sm sm:text-base">
                    Lengkapi Biodata Buku Induk — {editStudentTarget.namaSiswa}
                  </h3>
                  <p className="text-[11px] text-emerald-100">
                    Semua perubahan otomatis tersimpan &amp; tersinkron ke Data Siswa
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setEditModalOpen(false)}
                className="text-white/80 hover:text-white cursor-pointer"
              >
                <X size={18} />
              </button>
            </div>

            {/* Sub-tab Edit */}
            <div className="flex flex-wrap gap-1 px-4 pt-3 pb-2 bg-slate-50 border-b border-slate-200 text-xs shrink-0">
              {[
                { id: 'pribadi', label: 'A–B. Pribadi & Foto' },
                { id: 'kesehatan', label: 'C–D. Kesehatan & Pendidikan' },
                { id: 'ortu', label: 'E–G. Ayah, Ibu & Wali' },
                { id: 'perkembangan', label: 'H. Hobi & Kelulusan/Mutasi' }
              ].map((t) => (
                <button
                  key={t.id}
                  type="button"
                  onClick={() => setActiveEditSection(t.id as typeof activeEditSection)}
                  className={`px-3 py-1.5 rounded-lg font-bold cursor-pointer ${
                    activeEditSection === t.id
                      ? 'bg-emerald-700 text-white'
                      : 'bg-white text-slate-700 border border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  {t.label}
                </button>
              ))}
            </div>

            <form onSubmit={handleSaveBukuInduk} className="p-4 sm:p-5 overflow-y-auto space-y-4 text-xs flex-1">
              {activeEditSection === 'pribadi' && (
                <div className="space-y-3">
                  {/* Pas Foto */}
                  <div className="flex items-center gap-4 p-3 rounded-xl bg-slate-50 border border-slate-200">
                    <div className="w-16 h-20 border-2 border-slate-400 bg-white flex items-center justify-center overflow-hidden shrink-0">
                      {formBase.fotoUrl ? (
                        <img src={formBase.fotoUrl} alt="Foto Siswa" className="w-full h-full object-cover" />
                      ) : (
                        <span className="text-[10px] text-slate-400 text-center">3×4</span>
                      )}
                    </div>
                    <div className="space-y-1.5">
                      <div className="font-bold text-slate-800">Pas Foto Siswa (3 × 4 cm)</div>
                      <div className="flex flex-wrap items-center gap-2">
                        <input
                          ref={photoInputRef}
                          type="file"
                          accept="image/*"
                          onChange={handlePhotoUpload}
                          className="hidden"
                        />
                        <button
                          type="button"
                          onClick={() => photoInputRef.current?.click()}
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold cursor-pointer"
                        >
                          <Upload size={13} /> Unggah Foto
                        </button>
                        {formBase.fotoUrl && (
                          <button
                            type="button"
                            onClick={() => setFormBase((p) => ({ ...p, fotoUrl: '' }))}
                            className="px-2.5 py-1.5 bg-rose-50 text-rose-700 border border-rose-200 rounded-lg font-bold cursor-pointer"
                          >
                            Hapus Foto
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div className="sm:col-span-2">
                      <label className="block font-semibold mb-1">Nama Lengkap *</label>
                      <input
                        type="text"
                        required
                        value={formBase.namaSiswa || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, namaSiswa: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg font-bold"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Nama Panggilan</label>
                      <input
                        type="text"
                        value={formBi.namaPanggilan || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, namaPanggilan: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">NIS (No. Induk)</label>
                      <input
                        type="text"
                        value={formBase.nis || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, nis: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg font-mono"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">NISN</label>
                      <input
                        type="text"
                        value={formBase.nisn || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, nisn: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg font-mono"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Jenis Kelamin</label>
                      <select
                        value={formBase.jenisKelamin || 'L'}
                        onChange={(e) => setFormBase((p) => ({ ...p, jenisKelamin: e.target.value as 'L' | 'P' }))}
                        className="w-full px-3 py-1.5 border rounded-lg"
                      >
                        <option value="L">Laki-laki</option>
                        <option value="P">Perempuan</option>
                      </select>
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Tempat Lahir</label>
                      <input
                        type="text"
                        value={formBase.tempatLahir || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, tempatLahir: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Tanggal Lahir</label>
                      <input
                        type="date"
                        value={formBase.tanggalLahir || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, tanggalLahir: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Kewarganegaraan</label>
                      <input
                        type="text"
                        value={formBi.kewarganegaraan || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, kewarganegaraan: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Anak ke-</label>
                      <input
                        type="text"
                        value={formBi.anakKe || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, anakKe: e.target.value }))}
                        placeholder="Contoh: 2"
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Jml Saudara Kandung / Tiri / Angkat</label>
                      <div className="grid grid-cols-3 gap-1">
                        <input
                          type="text"
                          value={formBi.jumlahSaudaraKandung || ''}
                          onChange={(e) => setFormBi((p) => ({ ...p, jumlahSaudaraKandung: e.target.value }))}
                          placeholder="Kndg"
                          className="w-full px-2 py-1.5 border rounded-lg"
                        />
                        <input
                          type="text"
                          value={formBi.jumlahSaudaraTiri || ''}
                          onChange={(e) => setFormBi((p) => ({ ...p, jumlahSaudaraTiri: e.target.value }))}
                          placeholder="Tiri"
                          className="w-full px-2 py-1.5 border rounded-lg"
                        />
                        <input
                          type="text"
                          value={formBi.jumlahSaudaraAngkat || ''}
                          onChange={(e) => setFormBi((p) => ({ ...p, jumlahSaudaraAngkat: e.target.value }))}
                          placeholder="Angkt"
                          className="w-full px-2 py-1.5 border rounded-lg"
                        />
                      </div>
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Status Anak</label>
                      <select
                        value={formBi.anakYatim || 'Lengkap'}
                        onChange={(e) => setFormBi((p) => ({ ...p, anakYatim: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg"
                      >
                        <option value="Lengkap">Lengkap (Ayah &amp; Ibu)</option>
                        <option value="Yatim">Yatim</option>
                        <option value="Piatu">Piatu</option>
                        <option value="Yatim Piatu">Yatim Piatu</option>
                      </select>
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Bahasa Sehari-hari</label>
                      <input
                        type="text"
                        value={formBi.bahasaSehariHari || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, bahasaSehariHari: e.target.value }))}
                        placeholder="Contoh: Bahasa Indonesia / Daerah"
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Nomor Telepon / HP</label>
                      <input
                        type="text"
                        value={formBase.telepon || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, telepon: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Tinggal Dengan</label>
                      <input
                        type="text"
                        value={formBi.tinggalDengan || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, tinggalDengan: e.target.value }))}
                        placeholder="Orang Tua / Wali / Asrama"
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div className="sm:col-span-2">
                      <label className="block font-semibold mb-1">Alamat Lengkap Tempat Tinggal</label>
                      <input
                        type="text"
                        value={formBase.alamat || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, alamat: e.target.value }))}
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Jarak / Transportasi ke Sekolah</label>
                      <input
                        type="text"
                        value={formBi.jarakKeSekolah || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, jarakKeSekolah: e.target.value }))}
                        placeholder="Contoh: 1 km / Jalan kaki"
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                  </div>
                </div>
              )}

              {activeEditSection === 'kesehatan' && (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block font-semibold mb-1">Golongan Darah</label>
                    <select
                      value={formBi.golonganDarah || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, golonganDarah: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    >
                      <option value="">Belum diketahui / -</option>
                      <option value="A">A</option>
                      <option value="B">B</option>
                      <option value="AB">AB</option>
                      <option value="O">O</option>
                    </select>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="block font-semibold mb-1">Tinggi Badan (cm)</label>
                      <input
                        type="text"
                        value={formBi.tinggiBadan || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, tinggiBadan: e.target.value }))}
                        placeholder="135"
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                    <div>
                      <label className="block font-semibold mb-1">Berat Badan (kg)</label>
                      <input
                        type="text"
                        value={formBi.beratBadan || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, beratBadan: e.target.value }))}
                        placeholder="30"
                        className="w-full px-3 py-1.5 border rounded-lg"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Penyakit yang Pernah Diderita</label>
                    <input
                      type="text"
                      value={formBi.penyakitPernahDiderita || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, penyakitPernahDiderita: e.target.value }))}
                      placeholder="Tidak ada / Malaria / dll"
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Kelainan Jasmani</label>
                    <input
                      type="text"
                      value={formBi.kelainanJasmani || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, kelainanJasmani: e.target.value }))}
                      placeholder="Tidak ada"
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Pendidikan Sebelumnya (TK/PAUD/Asal Sekolah)</label>
                    <input
                      type="text"
                      value={formBi.asalTkPaud || formBase.sekolahAsal || ''}
                      onChange={(e) => {
                        setFormBi((p) => ({ ...p, asalTkPaud: e.target.value }));
                        setFormBase((p) => ({ ...p, sekolahAsal: e.target.value }));
                      }}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Tanggal &amp; No. STTB / Surat Keterangan</label>
                    <input
                      type="text"
                      value={formBi.tanggalNoSttbSebelumnya || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, tanggalNoSttbSebelumnya: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Diterima di Kelas</label>
                    <input
                      type="text"
                      value={formBase.diterimaKelas || ''}
                      onChange={(e) => setFormBase((p) => ({ ...p, diterimaKelas: e.target.value }))}
                      placeholder="Contoh: Kelas I"
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Tanggal Diterima di Sekolah Ini</label>
                    <input
                      type="date"
                      value={formBase.diterimaTanggal || ''}
                      onChange={(e) => setFormBase((p) => ({ ...p, diterimaTanggal: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Pindahan Dari Sekolah (Jika Pindahan)</label>
                    <input
                      type="text"
                      value={formBi.pindahanDariSekolah || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, pindahanDariSekolah: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Alasan Pindah</label>
                    <input
                      type="text"
                      value={formBi.alasanPindah || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, alasanPindah: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                </div>
              )}

              {activeEditSection === 'ortu' && (
                <div className="space-y-4">
                  <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
                    <div className="font-bold uppercase text-slate-700">E. Ayah Kandung</div>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                      <input
                        type="text"
                        placeholder="Nama Ayah"
                        value={formBase.namaAyah || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, namaAyah: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white font-semibold"
                      />
                      <input
                        type="text"
                        placeholder="Tempat, Tgl Lahir Ayah"
                        value={formBi.ttlAyah || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, ttlAyah: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Pendidikan Terakhir"
                        value={formBi.pendidikanAyah || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, pendidikanAyah: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Pekerjaan Ayah"
                        value={formBase.pekerjaanAyah || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, pekerjaanAyah: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Penghasilan / Bulan"
                        value={formBi.penghasilanAyah || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, penghasilanAyah: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <select
                        value={formBi.statusAyah || 'Masih Hidup'}
                        onChange={(e) => setFormBi((p) => ({ ...p, statusAyah: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      >
                        <option value="Masih Hidup">Masih Hidup</option>
                        <option value="Meninggal Dunia">Meninggal Dunia</option>
                      </select>
                    </div>
                  </div>

                  <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
                    <div className="font-bold uppercase text-slate-700">F. Ibu Kandung</div>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                      <input
                        type="text"
                        placeholder="Nama Ibu"
                        value={formBase.namaIbu || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, namaIbu: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white font-semibold"
                      />
                      <input
                        type="text"
                        placeholder="Tempat, Tgl Lahir Ibu"
                        value={formBi.ttlIbu || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, ttlIbu: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Pendidikan Terakhir"
                        value={formBi.pendidikanIbu || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, pendidikanIbu: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Pekerjaan Ibu"
                        value={formBase.pekerjaanIbu || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, pekerjaanIbu: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Penghasilan / Bulan"
                        value={formBi.penghasilanIbu || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, penghasilanIbu: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <select
                        value={formBi.statusIbu || 'Masih Hidup'}
                        onChange={(e) => setFormBi((p) => ({ ...p, statusIbu: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      >
                        <option value="Masih Hidup">Masih Hidup</option>
                        <option value="Meninggal Dunia">Meninggal Dunia</option>
                      </select>
                    </div>
                  </div>

                  <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
                    <div className="font-bold uppercase text-slate-700">G. Wali Peserta Didik (Opsional)</div>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                      <input
                        type="text"
                        placeholder="Nama Wali"
                        value={formBase.namaWali || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, namaWali: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Hubungan Keluarga"
                        value={formBi.hubunganWali || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, hubunganWali: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Pekerjaan Wali"
                        value={formBase.pekerjaanWali || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, pekerjaanWali: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Pendidikan Wali"
                        value={formBi.pendidikanWali || ''}
                        onChange={(e) => setFormBi((p) => ({ ...p, pendidikanWali: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white"
                      />
                      <input
                        type="text"
                        placeholder="Alamat Wali"
                        value={formBase.alamatWali || ''}
                        onChange={(e) => setFormBase((p) => ({ ...p, alamatWali: e.target.value }))}
                        className="px-2.5 py-1.5 border rounded-lg bg-white sm:col-span-2"
                      />
                    </div>
                  </div>
                </div>
              )}

              {activeEditSection === 'perkembangan' && (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block font-semibold mb-1">Kegemaran Kesenian</label>
                    <input
                      type="text"
                      value={formBi.hobiKesenian || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, hobiKesenian: e.target.value }))}
                      placeholder="Menyanyi, Menggambar, Tari"
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Kegemaran Olahraga</label>
                    <input
                      type="text"
                      value={formBi.hobiOlahraga || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, hobiOlahraga: e.target.value }))}
                      placeholder="Sepak Bola, Atletik, Voli"
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Kemasyarakatan / Organisasi</label>
                    <input
                      type="text"
                      value={formBi.hobiOrganisasi || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, hobiOrganisasi: e.target.value }))}
                      placeholder="Pramuka, UKS"
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Beasiswa / Penghargaan</label>
                    <input
                      type="text"
                      value={formBi.beasiswa || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, beasiswa: e.target.value }))}
                      placeholder="PIP / Prestasi Akademik"
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Tanggal Tamat Belajar (Lulus)</label>
                    <input
                      type="date"
                      value={formBi.tamatBelajarTanggal || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, tamatBelajarTanggal: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Nomor Ijazah / STTB</label>
                    <input
                      type="text"
                      value={formBi.noIjazahSttb || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, noIjazahSttb: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg font-mono"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Tanggal Meninggalkan / Pindah Sekolah</label>
                    <input
                      type="date"
                      value={formBi.tanggalMeninggalkan || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, tanggalMeninggalkan: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Alasan Meninggalkan Sekolah</label>
                    <input
                      type="text"
                      value={formBi.alasanMeninggalkan || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, alasanMeninggalkan: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Melanjutkan ke Sekolah</label>
                    <input
                      type="text"
                      value={formBi.melanjutkanKe || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, melanjutkanKe: e.target.value }))}
                      placeholder="SMP Negeri ..."
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                  <div>
                    <label className="block font-semibold mb-1">Catatan Khusus Perkembangan</label>
                    <input
                      type="text"
                      value={formBi.catatanKhusus || ''}
                      onChange={(e) => setFormBi((p) => ({ ...p, catatanKhusus: e.target.value }))}
                      className="w-full px-3 py-1.5 border rounded-lg"
                    />
                  </div>
                </div>
              )}

              <div className="pt-3 border-t flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setEditModalOpen(false)}
                  className="px-4 py-2 border border-slate-300 rounded-xl font-semibold text-slate-700 hover:bg-slate-50 cursor-pointer"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded-xl font-bold inline-flex items-center gap-1.5 cursor-pointer"
                >
                  <Save size={14} /> Simpan Buku Induk
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
