import React, { useEffect, useRef, useState } from 'react';
import {
  Home,
  HelpCircle,
  Trash2,
  Save,
  UserPlus,
  Search,
  Filter,
  Download,
  FileSpreadsheet,
  FileUp,
  FolderDown,
  Link2,
  BookUser,
  Eye,
  Edit2,
  CheckCircle2,
  AlertCircle,
  X,
  GraduationCap
} from 'lucide-react';
import { useApp, getFaseFromKelas } from '../../context/AppContext';
import { Student, ClassLevel, FaseType, ALL_CLASSES } from '../../types';
import { SaveStatusIndicator } from '../common/SaveStatusIndicator';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { useSaveFeedback } from '../../hooks/useSaveFeedback';
import {
  exportStudentsToExcel,
  exportStudentsToCsv,
  exportStudentTemplateToExcel
} from '../../utils/exportHelper';
import { parseAndImportTemplate } from '../../utils/templateImporter';

const AGAMA_OPTIONS = ['Kristen', 'Katolik', 'Islam', 'Hindu', 'Buddha', 'Konghucu'] as const;
const MIN_SPREADSHEET_ROWS = 60;

interface SpreadsheetDraftRow {
  rowIndex: number;
  existingId?: string;
  nis: string;
  nisn: string;
  namaSiswa: string;
  namaPanggilan: string;
  ttl: string;
  jenisKelamin: 'L' | 'P' | '';
  agama: string;
  sekolahAsal: string;
  alamat: string;
  namaAyah: string;
  namaIbu: string;
  pekerjaanAyah: string;
  pekerjaanIbu: string;
  alamatOrtuJalan: string;
  alamatOrtuDesa: string;
  alamatOrtuKecamatan: string;
  alamatOrtuKabupaten: string;
  alamatOrtuProvinsi: string;
  namaWali: string;
  pekerjaanWali: string;
  alamatWali: string;
}

const cleanDisplay = (val?: string): string => {
  if (!val || val === '-' || val === '_') return '';
  return val;
};

const formatTtlCell = (tempat?: string, tanggal?: string): string => {
  const t = cleanDisplay(tempat);
  const d = cleanDisplay(tanggal);
  if (t && d) return `${t}, ${d}`;
  return t || d;
};

const parseTtlCell = (ttl: string): { tempatLahir: string; tanggalLahir: string } => {
  const trimmed = ttl.trim();
  if (!trimmed) return { tempatLahir: '-', tanggalLahir: '-' };
  if (trimmed.includes(',')) {
    const [tp, ...tg] = trimmed.split(',');
    return {
      tempatLahir: tp.trim() || '-',
      tanggalLahir: tg.join(',').trim() || '-'
    };
  }
  return { tempatLahir: trimmed, tanggalLahir: '-' };
};

const studentToDraftRow = (s: Student, rowIndex: number): SpreadsheetDraftRow => ({
  rowIndex,
  existingId: s.idSiswa,
  nis: cleanDisplay(s.nis),
  nisn: s.nisn && !s.nisn.startsWith('BLM_') ? cleanDisplay(s.nisn) : '',
  namaSiswa: s.namaSiswa || '',
  namaPanggilan: cleanDisplay(s.namaPanggilan || s.bukuInduk?.namaPanggilan),
  ttl: formatTtlCell(s.tempatLahir, s.tanggalLahir),
  jenisKelamin: s.jenisKelamin || 'L',
  agama: cleanDisplay(s.agama),
  sekolahAsal: cleanDisplay(s.sekolahAsal || s.bukuInduk?.asalTkPaud),
  alamat: cleanDisplay(s.alamat),
  namaAyah: cleanDisplay(s.namaAyah),
  namaIbu: cleanDisplay(s.namaIbu),
  pekerjaanAyah: cleanDisplay(s.pekerjaanAyah),
  pekerjaanIbu: cleanDisplay(s.pekerjaanIbu),
  alamatOrtuJalan: cleanDisplay(s.alamatOrtuJalan),
  alamatOrtuDesa: cleanDisplay(s.alamatOrtuDesa),
  alamatOrtuKecamatan: cleanDisplay(s.alamatOrtuKecamatan),
  alamatOrtuKabupaten: cleanDisplay(s.alamatOrtuKabupaten),
  alamatOrtuProvinsi: cleanDisplay(s.alamatOrtuProvinsi),
  namaWali: cleanDisplay(s.namaWali),
  pekerjaanWali: cleanDisplay(s.pekerjaanWali),
  alamatWali: cleanDisplay(s.alamatWali)
});

const emptyDraftRow = (rowIndex: number): SpreadsheetDraftRow => ({
  rowIndex,
  nis: '',
  nisn: '',
  namaSiswa: '',
  namaPanggilan: '',
  ttl: '',
  jenisKelamin: '',
  agama: '',
  sekolahAsal: '',
  alamat: '',
  namaAyah: '',
  namaIbu: '',
  pekerjaanAyah: '',
  pekerjaanIbu: '',
  alamatOrtuJalan: '',
  alamatOrtuDesa: '',
  alamatOrtuKecamatan: '',
  alamatOrtuKabupaten: '',
  alamatOrtuProvinsi: '',
  namaWali: '',
  pekerjaanWali: '',
  alamatWali: ''
});

export const StudentSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    saveSiswa,
    saveSiswaBatch,
    deleteSiswa,
    saveAllData,
    syncAllStudentConnections,
    selectedKelas,
    setSelectedKelas,
    selectedSemester,
    filteredStudent,
    setSelectedStudentFilter,
    setActiveMenu,
    setPrintTargetStudentId,
    openReportPreview,
    openClearDataModal,
    showSaveNotification
  } = appContext;

  const { status: modalSaveStatus, triggerSave: triggerModalSave } = useSaveFeedback();
  const { status: toolbarSaveStatus, triggerSave: triggerToolbarSave } = useSaveFeedback();

  const [search, setSearch] = useState('');
  const [filterKelas, setFilterKelas] = useState<string>(selectedKelas);
  const [selectedStudentId, setSelectedStudentId] = useState<string | null>(null);
  const [showHelp, setShowHelp] = useState(false);
  const [showAgamaGuide, setShowAgamaGuide] = useState(false);
  const [showOnlyFilled, setShowOnlyFilled] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [importBanner, setImportBanner] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [deleteConfirmTarget, setDeleteConfirmTarget] = useState<{ id: string; name: string } | null>(null);

  const [modalOpen, setModalOpen] = useState(false);
  const [editingSiswa, setEditingSiswa] = useState<Student | null>(null);
  const directImportRef = useRef<HTMLInputElement>(null);

  const getFase = (kelas: ClassLevel): FaseType => getFaseFromKelas(kelas);

  const activeTargetKelas: ClassLevel =
    filterKelas === 'SEMUA' ? selectedKelas : (filterKelas as ClassLevel);

  const filteredStudents = db.siswa.filter((s) => {
    if (filteredStudent && s.idSiswa !== filteredStudent.idSiswa) {
      return false;
    }
    const q = search.trim().toLowerCase();
    const matchesKelas = filterKelas === 'SEMUA' || s.kelas === filterKelas;
    if (!q) return matchesKelas;
    const matchesSearch =
      s.namaSiswa.toLowerCase().includes(q) ||
      (s.namaPanggilan || s.bukuInduk?.namaPanggilan || '').toLowerCase().includes(q) ||
      s.nisn.toLowerCase().includes(q) ||
      s.nis.toLowerCase().includes(q) ||
      (s.alamat || '').toLowerCase().includes(q) ||
      (s.namaAyah || '').toLowerCase().includes(q) ||
      (s.namaIbu || '').toLowerCase().includes(q) ||
      (s.namaWali || '').toLowerCase().includes(q);
    return matchesKelas && matchesSearch;
  });

  const targetRowCount = showOnlyFilled
    ? Math.max(filteredStudents.length, 1)
    : Math.max(
        MIN_SPREADSHEET_ROWS,
        db.sekolah.jumlahSiswaProfil || 0,
        filteredStudents.length
      );

  const [draftRows, setDraftRows] = useState<SpreadsheetDraftRow[]>([]);

  useEffect(() => {
    const nextRows: SpreadsheetDraftRow[] = Array.from({ length: targetRowCount }, (_, idx) => {
      const student = filteredStudents[idx];
      return student ? studentToDraftRow(student, idx) : emptyDraftRow(idx);
    });
    setDraftRows(nextRows);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [db.siswa, filterKelas, search, filteredStudent?.idSiswa, targetRowCount]);

  const buildStudentFromDraft = (row: SpreadsheetDraftRow): Student | null => {
    const cleanName = row.namaSiswa.trim();
    if (!cleanName) return null;

    const existing = row.existingId
      ? db.siswa.find((s) => s.idSiswa === row.existingId)
      : undefined;

    const baseId = existing?.idSiswa || `S_${Date.now()}_${row.rowIndex + 1}`;
    const rawNisn = row.nisn.trim();
    const safeNisn =
      rawNisn && rawNisn !== '-'
        ? rawNisn
        : existing?.nisn && existing.nisn.startsWith('BLM_')
        ? existing.nisn
        : `BLM_${baseId}`;

    const { tempatLahir, tanggalLahir } = parseTtlCell(row.ttl);
    const kelas = existing?.kelas || activeTargetKelas;
    const fase = existing?.fase || getFase(kelas);
    const panggilan = row.namaPanggilan.trim();
    const asalSekolah = row.sekolahAsal.trim();

    return {
      ...(existing || {}),
      idSiswa: baseId,
      nisn: safeNisn,
      nis: row.nis.trim() || '-',
      namaSiswa: cleanName,
      namaPanggilan: panggilan,
      jenisKelamin: row.jenisKelamin === 'P' ? 'P' : 'L',
      tempatLahir,
      tanggalLahir,
      nik: existing?.nik || '-',
      agama: row.agama || existing?.agama || 'Kristen',
      kelas,
      fase,
      semester: existing?.semester || selectedSemester,
      tahunPelajaran: existing?.tahunPelajaran || db.sekolah.tahunPelajaran,
      namaAyah: row.namaAyah.trim() || '-',
      namaIbu: row.namaIbu.trim() || '-',
      pekerjaanAyah: row.pekerjaanAyah.trim(),
      pekerjaanIbu: row.pekerjaanIbu.trim(),
      alamatOrtuJalan: row.alamatOrtuJalan.trim(),
      alamatOrtuDesa: row.alamatOrtuDesa.trim(),
      alamatOrtuKecamatan: row.alamatOrtuKecamatan.trim(),
      alamatOrtuKabupaten: row.alamatOrtuKabupaten.trim(),
      alamatOrtuProvinsi: row.alamatOrtuProvinsi.trim(),
      namaWali: row.namaWali.trim(),
      pekerjaanWali: row.pekerjaanWali.trim(),
      alamatWali: row.alamatWali.trim(),
      sekolahAsal: asalSekolah,
      alamat: row.alamat.trim() || '-',
      statusSiswa: existing?.statusSiswa || 'Aktif',
      bukuInduk: {
        ...(existing?.bukuInduk || {}),
        namaPanggilan: panggilan,
        asalTkPaud: asalSekolah
      }
    };
  };

  const handleDraftCellChange = (
    rowIndex: number,
    field: keyof SpreadsheetDraftRow,
    value: string,
    commitImmediately = false
  ) => {
    setDraftRows((prev) => {
      const next = [...prev];
      const current = next[rowIndex];
      if (!current) return prev;
      const updated: SpreadsheetDraftRow = {
        ...current,
        [field]: value
      };
      next[rowIndex] = updated;

      if (commitImmediately && updated.namaSiswa.trim()) {
        const studentObj = buildStudentFromDraft(updated);
        if (studentObj) {
          const existing = updated.existingId
            ? db.siswa.find((s) => s.idSiswa === updated.existingId)
            : undefined;
          saveSiswa(studentObj, existing?.nisn);
        }
      }
      return next;
    });
  };

  const handleRowBlurCommit = (rowIndex: number) => {
    const row = draftRows[rowIndex];
    if (!row || !row.namaSiswa.trim()) return;
    const studentObj = buildStudentFromDraft(row);
    if (!studentObj) return;
    const existing = row.existingId ? db.siswa.find((s) => s.idSiswa === row.existingId) : undefined;
    saveSiswa(studentObj, existing?.nisn);
  };

  const handleSaveAllSpreadsheet = () => {
    const rowsToSave = draftRows
      .map((r) => buildStudentFromDraft(r))
      .filter((s): s is Student => s !== null);

    triggerToolbarSave(async () => {
      if (rowsToSave.length > 0) {
        saveSiswaBatch(rowsToSave);
      }
      syncAllStudentConnections();
      saveAllData();
      setImportBanner({
        type: 'success',
        text: `Seluruh Identitas Siswa (${rowsToSave.length} baris terisi) berhasil disimpan & dihubungkan ke Rapor Akademik, PTS, P5, Buku Induk, dan Piagam.`
      });
      showSaveNotification('Identitas Siswa berhasil disimpan & terhubung otomatis!');
    });
  };

  // Modal Form state for detailed editing/adding
  const defaultNewSiswa: Student = {
    idSiswa: `S_${Date.now()}`,
    nisn: '',
    nis: '',
    namaSiswa: '',
    namaPanggilan: '',
    jenisKelamin: 'L',
    tempatLahir: '',
    tanggalLahir: '',
    nik: '',
    agama: 'Kristen',
    kelas: activeTargetKelas || 'I',
    fase: getFase(activeTargetKelas || 'I'),
    semester: selectedSemester,
    tahunPelajaran: db.sekolah.tahunPelajaran,
    namaAyah: '',
    namaIbu: '',
    pekerjaanAyah: '',
    pekerjaanIbu: '',
    alamatOrtuJalan: '',
    alamatOrtuDesa: '',
    alamatOrtuKecamatan: '',
    alamatOrtuKabupaten: '',
    alamatOrtuProvinsi: '',
    namaWali: '',
    pekerjaanWali: '',
    alamatWali: '',
    sekolahAsal: '',
    alamat: '',
    statusSiswa: 'Aktif'
  };

  const [formData, setFormData] = useState<Student>(defaultNewSiswa);

  const handleOpenAdd = () => {
    const activeKls = activeTargetKelas || 'I';
    setEditingSiswa(null);
    setFormData({
      ...defaultNewSiswa,
      idSiswa: `S_${Date.now()}`,
      kelas: activeKls,
      fase: getFase(activeKls)
    });
    setModalOpen(true);
  };

  const handleOpenEdit = (siswa: Student) => {
    setEditingSiswa(siswa);
    setFormData({
      ...siswa,
      nisn: siswa.nisn.startsWith('BLM_') ? '' : cleanDisplay(siswa.nisn),
      nis: cleanDisplay(siswa.nis),
      namaPanggilan: cleanDisplay(siswa.namaPanggilan || siswa.bukuInduk?.namaPanggilan),
      tempatLahir: cleanDisplay(siswa.tempatLahir),
      tanggalLahir: cleanDisplay(siswa.tanggalLahir),
      nik: cleanDisplay(siswa.nik),
      sekolahAsal: cleanDisplay(siswa.sekolahAsal || siswa.bukuInduk?.asalTkPaud),
      pekerjaanAyah: cleanDisplay(siswa.pekerjaanAyah),
      pekerjaanIbu: cleanDisplay(siswa.pekerjaanIbu),
      alamatOrtuJalan: cleanDisplay(siswa.alamatOrtuJalan),
      alamatOrtuDesa: cleanDisplay(siswa.alamatOrtuDesa),
      alamatOrtuKecamatan: cleanDisplay(siswa.alamatOrtuKecamatan),
      alamatOrtuKabupaten: cleanDisplay(siswa.alamatOrtuKabupaten),
      alamatOrtuProvinsi: cleanDisplay(siswa.alamatOrtuProvinsi),
      namaWali: cleanDisplay(siswa.namaWali),
      pekerjaanWali: cleanDisplay(siswa.pekerjaanWali),
      alamatWali: cleanDisplay(siswa.alamatWali)
    });
    setModalOpen(true);
  };

  const handleModalSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const rawNisn = formData.nisn ? formData.nisn.trim() : '';
    const safeNisn = rawNisn === '' || rawNisn === '-' ? `BLM_${formData.idSiswa}` : rawNisn;
    const panggilan = (formData.namaPanggilan || '').trim();
    const sekolahAsal = (formData.sekolahAsal || '').trim();

    const cleanedSiswa: Student = {
      ...formData,
      nisn: safeNisn,
      nis: formData.nis ? formData.nis.trim() || '-' : '-',
      namaPanggilan: panggilan,
      tempatLahir: formData.tempatLahir ? formData.tempatLahir.trim() || '-' : '-',
      tanggalLahir: formData.tanggalLahir || '-',
      nik: formData.nik ? formData.nik.trim() || '-' : '-',
      sekolahAsal,
      pekerjaanAyah: formData.pekerjaanAyah ? formData.pekerjaanAyah.trim() : '',
      pekerjaanIbu: formData.pekerjaanIbu ? formData.pekerjaanIbu.trim() : '',
      alamatOrtuJalan: formData.alamatOrtuJalan ? formData.alamatOrtuJalan.trim() : '',
      alamatOrtuDesa: formData.alamatOrtuDesa ? formData.alamatOrtuDesa.trim() : '',
      alamatOrtuKecamatan: formData.alamatOrtuKecamatan ? formData.alamatOrtuKecamatan.trim() : '',
      alamatOrtuKabupaten: formData.alamatOrtuKabupaten ? formData.alamatOrtuKabupaten.trim() : '',
      alamatOrtuProvinsi: formData.alamatOrtuProvinsi ? formData.alamatOrtuProvinsi.trim() : '',
      namaWali: formData.namaWali ? formData.namaWali.trim() : '',
      pekerjaanWali: formData.pekerjaanWali ? formData.pekerjaanWali.trim() : '',
      alamatWali: formData.alamatWali ? formData.alamatWali.trim() : '',
      bukuInduk: {
        ...(formData.bukuInduk || {}),
        namaPanggilan: panggilan,
        asalTkPaud: sekolahAsal
      }
    };

    triggerModalSave(async () => {
      saveSiswa(cleanedSiswa, editingSiswa?.nisn);
      setTimeout(() => setModalOpen(false), 350);
    });
  };

  const handleImportFile = async (file: File) => {
    setIsImporting(true);
    setImportBanner(null);
    try {
      const result = await parseAndImportTemplate(file, 'template_siswa', appContext);
      if (result.success) {
        setSearch('');
        if (result.importedClasses && result.importedClasses.length === 1) {
          const targetKls = result.importedClasses[0];
          setFilterKelas(targetKls);
          setSelectedKelas(targetKls);
        } else {
          setFilterKelas('SEMUA');
        }
        saveAllData();
        setImportBanner({ type: 'success', text: result.message });
        showSaveNotification(result.message);
      } else {
        setImportBanner({ type: 'error', text: result.message });
        showSaveNotification(`Gagal impor: ${result.message}`);
      }
    } catch (e: unknown) {
      const msg = `Gagal membaca file Excel/CSV: ${
        e instanceof Error ? e.message : 'Format berkas tidak sesuai'
      }`;
      setImportBanner({ type: 'error', text: msg });
      showSaveNotification(msg);
    } finally {
      setIsImporting(false);
    }
  };

  const handleDirectFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleImportFile(file);
    e.target.value = '';
  };

  const handleExportExcel = () => {
    exportStudentsToExcel(filteredStudents, db.sekolah, filterKelas);
    showSaveNotification(`Berhasil mengekspor ${filteredStudents.length} data siswa ke Excel (.xlsx)!`);
  };

  const handleExportCsv = () => {
    exportStudentsToCsv(filteredStudents, db.sekolah, filterKelas);
    showSaveNotification(`Berhasil mengekspor ${filteredStudents.length} data siswa ke CSV (.csv)!`);
  };

  const handleDownloadTemplate = () => {
    const sourceRows = filteredStudents.length > 0 ? filteredStudents : db.siswa;
    exportStudentTemplateToExcel(sourceRows, db.sekolah, filterKelas);
    showSaveNotification('Template Excel Data Siswa (.xlsx) berhasil diunduh!');
  };

  const rombelBanner =
    filterKelas === 'SEMUA'
      ? 'SEMUA KELAS'
      : `KELAS ${
          filterKelas === db.sekolah.kelas && db.sekolah.rombel
            ? db.sekolah.rombel
            : filterKelas
        }`;

  return (
    <div className="bg-[#00ff00] p-2 sm:p-3 rounded-xl border-2 border-red-600 shadow-md space-y-2.5">
      {/* Hidden input untuk Impor Excel */}
      <input
        ref={directImportRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleDirectFileInputChange}
        className="hidden"
      />

      {/* TOP HEADER ROW matching image_a048e597.png */}
      <div className="grid grid-cols-1 xl:grid-cols-12 gap-2 items-stretch">
        {/* Top-Left Black Banner: Home, ?, IDENTITAS SISWA / KELAS / SEKOLAH, Trash */}
        <div className="xl:col-span-4 bg-black border-2 border-red-600 px-3 py-2 flex items-center justify-between gap-2 text-white shadow-sm">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setActiveMenu('dashboard')}
              title="Kembali ke Dashboard"
              aria-label="Beranda"
              className="w-8 h-8 rounded-full bg-gradient-to-b from-sky-400 to-blue-700 hover:brightness-110 border border-sky-200 flex items-center justify-center shrink-0 cursor-pointer"
            >
              <Home size={16} />
            </button>
            <button
              type="button"
              onClick={() => setShowHelp((prev) => !prev)}
              title="Panduan Pengisian Identitas Siswa"
              aria-label="Panduan Pengisian"
              className="w-8 h-8 rounded-full bg-gradient-to-b from-purple-400 to-purple-800 hover:brightness-110 border border-purple-200 flex items-center justify-center shrink-0 cursor-pointer"
            >
              <HelpCircle size={17} />
            </button>
          </div>

          <div className="text-center leading-tight min-w-0 flex-1">
            <div className="text-sm sm:text-base font-black text-yellow-300 tracking-wide uppercase truncate">
              IDENTITAS SISWA
            </div>
            <div className="text-xs font-extrabold text-yellow-200 uppercase truncate">
              {rombelBanner}
            </div>
            <div className="text-[11px] font-extrabold text-yellow-300 uppercase truncate">
              {db.sekolah.namaSekolah || 'SD NEGERI BALIL'}
            </div>
          </div>

          <button
            type="button"
            onClick={openClearDataModal}
            title="Hapus / Kosongkan Data"
            aria-label="Hapus Data"
            className="w-8 h-8 rounded-full bg-gradient-to-b from-slate-100 to-slate-300 hover:bg-rose-50 border-2 border-red-600 text-red-700 flex items-center justify-center shrink-0 cursor-pointer"
          >
            <Trash2 size={15} />
          </button>
        </div>

        {/* Callout Penting Kolom Agama + Cara Input Agama matching image_a048e597.png */}
        <div className="xl:col-span-5 bg-white border-2 border-red-600 px-3 py-1.5 text-[10px] leading-snug text-slate-900 flex flex-col justify-between">
          <div>
            <span className="font-black text-red-600">PENTING ! </span>
            <span>Pada KOLOM AGAMA, Pengguna Wajib Input Agama Siswa dari </span>
            <span className="font-bold text-red-600">DROPDOWN LIST</span>
            <span> yang disediakan pada kolom cyan (</span>
            <span className="font-bold text-red-600">JANGAN ASAL KETIK/COPY-PASTE</span>
            <span>).</span>
          </div>
          <div className="mt-1 flex flex-wrap items-center justify-between gap-2 pt-1 border-t border-slate-200">
            <span className="text-[10px] text-slate-700">
              Jika agama belum dipilih, Nama Siswa &amp; deskripsi Rapor Agama tidak aktif.
            </span>
            <button
              type="button"
              onClick={() => setShowAgamaGuide((p) => !p)}
              className="px-2 py-0.5 bg-red-600 hover:bg-red-700 text-white font-bold rounded text-[10px] cursor-pointer shrink-0"
            >
              Cara Input Agama
            </button>
          </div>
        </div>

        {/* Action Strip */}
        <div className="xl:col-span-3 bg-white border-2 border-red-600 px-2.5 py-2 flex flex-wrap items-center justify-end gap-1.5">
          <button
            type="button"
            onClick={handleSaveAllSpreadsheet}
            disabled={toolbarSaveStatus === 'saving'}
            className="inline-flex items-center gap-1 px-2.5 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded-md text-[11px] font-bold cursor-pointer shadow-xs disabled:opacity-60"
            title="Simpan & Hubungkan Semua Data Siswa"
          >
            <Save size={13} />
            <span>Simpan Semua</span>
          </button>
          <button
            type="button"
            onClick={handleOpenAdd}
            className="inline-flex items-center gap-1 px-2 py-1.5 bg-blue-700 hover:bg-blue-800 text-white rounded-md text-[11px] font-bold cursor-pointer"
            title="Tambah Siswa via Form Lengkap"
          >
            <UserPlus size={13} />
            <span>Tambah</span>
          </button>
          <button
            type="button"
            onClick={() => directImportRef.current?.click()}
            disabled={isImporting}
            className="inline-flex items-center gap-1 px-2 py-1.5 bg-violet-700 hover:bg-violet-800 text-white rounded-md text-[11px] font-bold cursor-pointer disabled:opacity-60"
            title="Impor Excel (.xlsx)"
          >
            <FileUp size={13} />
            <span>Impor</span>
          </button>
          <button
            type="button"
            onClick={handleExportExcel}
            className="inline-flex items-center gap-1 px-2 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-md text-[11px] font-bold cursor-pointer"
            title="Expor Excel (.xlsx)"
          >
            <FileSpreadsheet size={13} />
            <span>.xlsx</span>
          </button>
          <button
            type="button"
            onClick={handleExportCsv}
            className="inline-flex items-center gap-1 px-2 py-1.5 bg-slate-700 hover:bg-slate-800 text-white rounded-md text-[11px] font-bold cursor-pointer"
            title="Expor CSV"
          >
            <Download size={13} />
            <span>CSV</span>
          </button>
          <button
            type="button"
            onClick={handleDownloadTemplate}
            className="inline-flex items-center gap-1 px-2 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-300 rounded-md text-[11px] font-bold cursor-pointer"
            title="Unduh Template Excel"
          >
            <FolderDown size={13} />
            <span>Template</span>
          </button>
          <SaveStatusIndicator status={toolbarSaveStatus} compact />
        </div>
      </div>

      {/* Cara Input Agama Popup/Panel */}
      {showAgamaGuide && (
        <div className="bg-white border-2 border-red-600 rounded-lg p-3 text-xs text-slate-800 flex items-start justify-between gap-3 shadow-xs">
          <div>
            <div className="font-bold text-red-700 mb-1">Cara Input Agama Siswa:</div>
            <ol className="list-decimal pl-5 space-y-0.5 text-[11px]">
              <li>Klik cell pada kolom <strong>Agama</strong> (latar cyan biru muda) pada baris siswa.</li>
              <li>Klik tanda panah dropdown yang muncul pada cell tersebut.</li>
              <li>Pilih agama siswa (<strong>Kristen, Katolik, Islam, Hindu, Buddha,</strong> atau <strong>Konghucu</strong>).</li>
            </ol>
          </div>
          <button
            type="button"
            onClick={() => setShowAgamaGuide(false)}
            className="p-1 text-slate-500 hover:text-slate-900 cursor-pointer"
            aria-label="Tutup panduan agama"
          >
            <X size={14} />
          </button>
        </div>
      )}

      {/* Help Drawer */}
      {showHelp && (
        <div className="bg-white border-2 border-purple-700 rounded-lg p-3 text-xs text-slate-800 flex items-start justify-between gap-3 shadow-xs">
          <div className="space-y-1">
            <div className="font-bold text-purple-900">Panduan Tabel Identitas Siswa:</div>
            <ul className="list-disc pl-5 space-y-0.5 text-[11px]">
              <li>
                Anda dapat mengetik langsung di setiap sel seperti Excel (NIS, NISN, Nama, Panggilan, TTL format <em>Tempat, Tanggal</em>, Pendidikan Sebelumnya, Alamat, Data Orang Tua, &amp; Wali). Perubahan otomatis tersimpan saat berpindah sel atau menekan <strong>Simpan Semua</strong>.
              </li>
              <li>
                Kolom <strong>Panggilan</strong> otomatis terhubung ke rumus <em>Deskripsi Rapor Diawali Nama Panggilan Siswa</em> dan halaman <strong>Buku Induk</strong>.
              </li>
              <li>
                Kolom <strong>Alamat Orang Tua</strong> (Jalan, Kelurahan/Desa, Kecamatan, Kabupaten/Kota, Provinsi) dan <strong>Pendidikan Sebelumnya</strong> langsung tercetak pada Halaman 3 Identitas Rapor.
              </li>
            </ul>
          </div>
          <button
            type="button"
            onClick={() => setShowHelp(false)}
            className="p-1 text-slate-500 hover:text-slate-900 cursor-pointer"
            aria-label="Tutup panduan"
          >
            <X size={14} />
          </button>
        </div>
      )}

      {/* Feedback Banner */}
      {importBanner && (
        <div
          className={`rounded-lg p-2.5 border-2 flex items-center justify-between gap-3 text-xs font-semibold bg-white ${
            importBanner.type === 'success'
              ? 'border-emerald-700 text-emerald-900'
              : 'border-rose-700 text-rose-900'
          }`}
        >
          <div className="flex items-center gap-2">
            {importBanner.type === 'success' ? (
              <CheckCircle2 size={16} className="text-emerald-700 shrink-0" />
            ) : (
              <AlertCircle size={16} className="text-rose-700 shrink-0" />
            )}
            <span>{importBanner.text}</span>
          </div>
          <button
            type="button"
            onClick={() => setImportBanner(null)}
            className="p-1 hover:bg-slate-100 rounded cursor-pointer"
            aria-label="Tutup notifikasi"
          >
            <X size={14} />
          </button>
        </div>
      )}

      {/* Filter, Pencarian & Hubungkan Siswa Bar */}
      <div className="bg-white border border-slate-800 rounded-lg px-3 py-2 flex flex-wrap items-center justify-between gap-2 text-xs">
        <div className="flex flex-wrap items-center gap-2">
          <div className="relative min-w-[190px]">
            <Search size={14} className="absolute left-2.5 top-2 text-slate-400" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Cari nama, NIS, NISN..."
              className="w-full pl-8 pr-2.5 py-1 text-xs border border-slate-300 rounded-md focus:ring-2 focus:ring-blue-500 focus:outline-hidden"
            />
          </div>

          <StudentFilterSelect classScope="all" compact />

          <div className="flex items-center gap-1 bg-slate-50 border border-slate-300 px-2 py-1 rounded-md">
            <Filter size={13} className="text-slate-600" />
            <select
              value={filterKelas}
              onChange={(e) => setFilterKelas(e.target.value)}
              aria-label="Filter kelas siswa"
              className="bg-transparent font-bold text-slate-800 focus:outline-hidden cursor-pointer text-xs"
            >
              <option value="SEMUA">Semua Kelas</option>
              {ALL_CLASSES.map((k) => (
                <option key={k} value={k}>
                  Kelas {k}
                </option>
              ))}
            </select>
          </div>

          <button
            type="button"
            onClick={() => setShowOnlyFilled((p) => !p)}
            className={`px-2.5 py-1 rounded-md border text-[11px] font-semibold cursor-pointer ${
              showOnlyFilled
                ? 'bg-blue-700 text-white border-blue-800'
                : 'bg-slate-100 text-slate-800 border-slate-300 hover:bg-slate-200'
            }`}
          >
            {showOnlyFilled
              ? `Tampilkan 60 Baris`
              : `Hanya Terisi (${filteredStudents.length})`}
          </button>
        </div>

        <div className="flex flex-wrap items-center gap-3 text-[11px] font-semibold text-slate-700">
          <span>
            Terisi: <strong className="text-slate-900">{filteredStudents.length}</strong> siswa
          </span>
          <span aria-hidden="true">·</span>
          <span>
            L: <strong className="text-blue-800">{filteredStudents.filter((s) => s.jenisKelamin === 'L').length}</strong>
          </span>
          <span>
            P: <strong className="text-pink-700">{filteredStudents.filter((s) => s.jenisKelamin === 'P').length}</strong>
          </span>
          <span aria-hidden="true">·</span>
          <span className="text-slate-600">
            Kr:{filteredStudents.filter((s) => s.agama === 'Kristen').length} ·
            Kt:{filteredStudents.filter((s) => s.agama === 'Katolik').length} ·
            Is:{filteredStudents.filter((s) => s.agama === 'Islam').length} ·
            Hn:{filteredStudents.filter((s) => s.agama === 'Hindu').length} ·
            Bd:{filteredStudents.filter((s) => s.agama === 'Buddha').length} ·
            Kh:{filteredStudents.filter((s) => s.agama === 'Konghucu' || s.agama === 'Khonghucu').length}
          </span>
        </div>
      </div>

      {/* EXCEL-STYLE DATA RAPOR / IDENTITAS SISWA SPREADSHEET */}
      <div className="bg-white border-2 border-black overflow-hidden shadow-sm">
        <div className="overflow-x-auto max-h-[68vh]">
          <table className="w-full border-collapse text-[11px] font-sans min-w-[2350px]">
            <thead className="sticky top-0 z-20">
              {/* Row 1: Purple Group Header Row */}
              <tr className="bg-[#5e35b1] text-yellow-300 font-bold text-xs">
                <th colSpan={4} className="border border-black bg-[#311b92] px-2 py-1 text-center">
                  IDENTITAS POKOK
                </th>
                <th colSpan={6} className="border border-black px-2 py-1 text-center tracking-[0.22em] text-sm">
                  DATA RAPOR
                </th>
                <th colSpan={2} className="border border-black px-2 py-1 text-center">
                  Nama Orang Tua
                </th>
                <th colSpan={2} className="border border-black px-2 py-1 text-center">
                  Pekerjaan Orang Tua
                </th>
                <th colSpan={5} className="border border-black px-2 py-1 text-center">
                  Alamat Orang Tua
                </th>
                <th colSpan={3} className="border border-black px-2 py-1 text-center">
                  Wali Peserta Didik
                </th>
                <th className="border border-black bg-[#311b92] px-2 py-1 text-center">
                  Terhubung
                </th>
              </tr>

              {/* Row 2: Blue Column Header Row */}
              <tr className="bg-[#1e40af] text-white font-bold text-[11px] whitespace-nowrap">
                <th className="border border-black px-2 py-1.5 w-10 text-center sticky left-0 z-30 bg-[#1e40af]">
                  No.
                </th>
                <th className="border border-black px-2 py-1.5 w-24 text-center">NIS</th>
                <th className="border border-black px-2 py-1.5 w-28 text-center">NISN</th>
                <th className="border border-black px-2 py-1.5 w-52 text-center sticky left-10 z-30 bg-[#1e40af]">
                  Nama
                </th>
                <th className="border border-black px-2 py-1.5 w-28 text-center">Panggilan</th>
                <th className="border border-black px-2 py-1.5 w-44 text-center">Tempat / Tanggal Lahir</th>
                <th className="border border-black px-2 py-1.5 w-24 text-center">Jenis Kelamin</th>
                <th className="border border-black px-2 py-1.5 w-28 text-center bg-[#00ffff] text-black border-t-2 border-t-red-600">
                  Agama
                </th>
                <th className="border border-black px-2 py-1.5 w-36 text-center">Pendidikan Sebelumnya</th>
                <th className="border border-black px-2 py-1.5 w-48 text-center">Alamat Peserta Didik</th>
                <th className="border border-black px-2 py-1.5 w-36 text-center">Ayah</th>
                <th className="border border-black px-2 py-1.5 w-36 text-center">Ibu</th>
                <th className="border border-black px-2 py-1.5 w-32 text-center">Ayah</th>
                <th className="border border-black px-2 py-1.5 w-32 text-center">Ibu</th>
                <th className="border border-black px-2 py-1.5 w-36 text-center">Jalan</th>
                <th className="border border-black px-2 py-1.5 w-32 text-center">Kelurahan/Desa</th>
                <th className="border border-black px-2 py-1.5 w-28 text-center">Kecamatan</th>
                <th className="border border-black px-2 py-1.5 w-32 text-center">Kabupaten/Kota</th>
                <th className="border border-black px-2 py-1.5 w-32 text-center">Provinsi</th>
                <th className="border border-black px-2 py-1.5 w-32 text-center">Nama</th>
                <th className="border border-black px-2 py-1.5 w-28 text-center">Pekerjaan</th>
                <th className="border border-black px-2 py-1.5 w-40 text-center">Alamat</th>
                <th className="border border-black px-2 py-1.5 w-36 text-center">Aksi</th>
              </tr>
            </thead>

            <tbody>
              {draftRows.map((row, idx) => {
                const linkedStudent = row.existingId
                  ? db.siswa.find((s) => s.idSiswa === row.existingId)
                  : undefined;
                const isSelected = Boolean(row.existingId && selectedStudentId === row.existingId);
                const isLinkedFilter = Boolean(
                  row.existingId && filteredStudent?.idSiswa === row.existingId
                );

                return (
                  <tr
                    key={row.existingId || `empty_${idx}`}
                    onClick={() => row.existingId && setSelectedStudentId(row.existingId)}
                    className={
                      isSelected
                        ? 'bg-blue-50'
                        : idx % 2 === 0
                        ? 'bg-white hover:bg-slate-50'
                        : 'bg-slate-50/40 hover:bg-slate-100/60'
                    }
                  >
                    {/* 1. No. */}
                    <td
                      className={`border border-black px-1.5 py-0.5 text-center font-mono font-bold text-slate-700 sticky left-0 z-10 ${
                        isSelected ? 'bg-blue-100' : 'bg-white'
                      }`}
                    >
                      {idx + 1}
                    </td>

                    {/* 2. NIS */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.nis}
                        onChange={(e) => handleDraftCellChange(idx, 'nis', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`NIS baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] font-mono bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 3. NISN */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.nisn}
                        onChange={(e) => handleDraftCellChange(idx, 'nisn', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`NISN baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] font-mono bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 4. Nama */}
                    <td
                      className={`border border-black p-0 sticky left-10 z-10 ${
                        isSelected ? 'bg-blue-50' : 'bg-white'
                      }`}
                    >
                      <input
                        type="text"
                        value={row.namaSiswa}
                        onChange={(e) => handleDraftCellChange(idx, 'namaSiswa', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        placeholder={idx === filteredStudents.length ? 'Ketik nama siswa baru...' : ''}
                        aria-label={`Nama siswa baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] font-bold text-slate-900 bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 5. Panggilan */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.namaPanggilan}
                        onChange={(e) => handleDraftCellChange(idx, 'namaPanggilan', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Panggilan baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 6. Tempat / Tanggal Lahir */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.ttl}
                        onChange={(e) => handleDraftCellChange(idx, 'ttl', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Tempat Tanggal Lahir baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 7. Jenis Kelamin */}
                    <td className="border border-black p-0">
                      <select
                        value={row.jenisKelamin}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'jenisKelamin', e.target.value, true)
                        }
                        aria-label={`Jenis Kelamin baris ${idx + 1}`}
                        className="w-full px-1 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden cursor-pointer text-center font-semibold"
                      >
                        <option value="">-</option>
                        <option value="L">L</option>
                        <option value="P">P</option>
                      </select>
                    </td>

                    {/* 8. Agama (Cyan column with mandatory dropdown) */}
                    <td className="border border-black p-0 bg-[#00ffff]">
                      <select
                        value={row.agama}
                        onChange={(e) => handleDraftCellChange(idx, 'agama', e.target.value, true)}
                        aria-label={`Agama baris ${idx + 1}`}
                        className="w-full px-1 py-1 text-[11px] font-bold text-black bg-[#00ffff] focus:bg-cyan-200 focus:outline-hidden cursor-pointer"
                      >
                        <option value="">Pilih...</option>
                        {AGAMA_OPTIONS.map((ag) => (
                          <option key={ag} value={ag}>
                            {ag}
                          </option>
                        ))}
                      </select>
                    </td>

                    {/* 9. Pendidikan Sebelumnya */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.sekolahAsal}
                        onChange={(e) => handleDraftCellChange(idx, 'sekolahAsal', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Pendidikan Sebelumnya baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 10. Alamat Peserta Didik */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.alamat}
                        onChange={(e) => handleDraftCellChange(idx, 'alamat', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Alamat Peserta Didik baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 11. Nama Ayah */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.namaAyah}
                        onChange={(e) => handleDraftCellChange(idx, 'namaAyah', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Nama Ayah baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 12. Nama Ibu */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.namaIbu}
                        onChange={(e) => handleDraftCellChange(idx, 'namaIbu', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Nama Ibu baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 13. Pekerjaan Ayah */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.pekerjaanAyah}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'pekerjaanAyah', e.target.value)
                        }
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Pekerjaan Ayah baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 14. Pekerjaan Ibu */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.pekerjaanIbu}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'pekerjaanIbu', e.target.value)
                        }
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Pekerjaan Ibu baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 15. Alamat Ortu - Jalan */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.alamatOrtuJalan}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'alamatOrtuJalan', e.target.value)
                        }
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Jalan Orang Tua baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 16. Alamat Ortu - Kelurahan/Desa */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.alamatOrtuDesa}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'alamatOrtuDesa', e.target.value)
                        }
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Kelurahan/Desa Orang Tua baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 17. Alamat Ortu - Kecamatan */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.alamatOrtuKecamatan}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'alamatOrtuKecamatan', e.target.value)
                        }
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Kecamatan Orang Tua baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 18. Alamat Ortu - Kabupaten/Kota */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.alamatOrtuKabupaten}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'alamatOrtuKabupaten', e.target.value)
                        }
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Kabupaten/Kota Orang Tua baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 19. Alamat Ortu - Provinsi */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.alamatOrtuProvinsi}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'alamatOrtuProvinsi', e.target.value)
                        }
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Provinsi Orang Tua baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 20. Wali - Nama */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.namaWali}
                        onChange={(e) => handleDraftCellChange(idx, 'namaWali', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Nama Wali baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 21. Wali - Pekerjaan */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.pekerjaanWali}
                        onChange={(e) =>
                          handleDraftCellChange(idx, 'pekerjaanWali', e.target.value)
                        }
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Pekerjaan Wali baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 22. Wali - Alamat */}
                    <td className="border border-black p-0">
                      <input
                        type="text"
                        value={row.alamatWali}
                        onChange={(e) => handleDraftCellChange(idx, 'alamatWali', e.target.value)}
                        onBlur={() => handleRowBlurCommit(idx)}
                        aria-label={`Alamat Wali baris ${idx + 1}`}
                        className="w-full px-1.5 py-1 text-[11px] bg-transparent focus:bg-yellow-50 focus:outline-hidden"
                      />
                    </td>

                    {/* 23. Aksi & Keterhubungan */}
                    <td className="border border-black px-1 py-0.5 text-center">
                      {linkedStudent ? (
                        <div className="flex items-center justify-center gap-1">
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              setSelectedStudentFilter(
                                isLinkedFilter ? 'SEMUA' : linkedStudent.idSiswa
                              );
                              showSaveNotification(
                                isLinkedFilter
                                  ? 'Filter Nama Siswa dilepas.'
                                  : `Nama siswa ${linkedStudent.namaSiswa} dihubungkan ke seluruh menu!`
                              );
                            }}
                            title={
                              isLinkedFilter
                                ? 'Lepas Hubungan Siswa'
                                : 'Hubungkan ke Semua Menu'
                            }
                            aria-label={`Hubungkan ${linkedStudent.namaSiswa}`}
                            className={`w-6 h-6 inline-flex items-center justify-center rounded border cursor-pointer ${
                              isLinkedFilter
                                ? 'bg-emerald-600 text-white border-emerald-700'
                                : 'bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100'
                            }`}
                          >
                            <Link2 size={12} />
                          </button>
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              setSelectedKelas(linkedStudent.kelas);
                              setPrintTargetStudentId(linkedStudent.idSiswa);
                              setActiveMenu('buku_induk');
                            }}
                            title="Buka Buku Induk"
                            aria-label={`Buku Induk ${linkedStudent.namaSiswa}`}
                            className="w-6 h-6 inline-flex items-center justify-center rounded bg-teal-50 text-teal-800 border border-teal-200 hover:bg-teal-100 cursor-pointer"
                          >
                            <BookUser size={12} />
                          </button>
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              openReportPreview(linkedStudent.idSiswa);
                            }}
                            title="Preview Rapor"
                            aria-label={`Preview Rapor ${linkedStudent.namaSiswa}`}
                            className="w-6 h-6 inline-flex items-center justify-center rounded bg-indigo-50 text-indigo-800 border border-indigo-200 hover:bg-indigo-100 cursor-pointer"
                          >
                            <Eye size={12} />
                          </button>
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              handleOpenEdit(linkedStudent);
                            }}
                            title="Edit Detail Lengkap"
                            aria-label={`Edit ${linkedStudent.namaSiswa}`}
                            className="w-6 h-6 inline-flex items-center justify-center rounded bg-blue-50 text-blue-800 border border-blue-200 hover:bg-blue-100 cursor-pointer"
                          >
                            <Edit2 size={12} />
                          </button>
                          <button
                            type="button"
                            onClick={(e) => {
                              e.stopPropagation();
                              setDeleteConfirmTarget({
                                id: linkedStudent.idSiswa,
                                name: linkedStudent.namaSiswa
                              });
                            }}
                            title="Hapus Siswa"
                            aria-label={`Hapus ${linkedStudent.namaSiswa}`}
                            className="w-6 h-6 inline-flex items-center justify-center rounded bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100 cursor-pointer"
                          >
                            <Trash2 size={12} />
                          </button>
                        </div>
                      ) : (
                        <span className="text-[10px] text-slate-400">-</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Inline Confirm Delete Modal */}
      {deleteConfirmTarget && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
          <div className="bg-white rounded-xl border-2 border-red-600 shadow-xl max-w-sm w-full p-5 space-y-4">
            <div className="text-sm font-bold text-slate-900">Hapus Data Siswa?</div>
            <p className="text-xs text-slate-600">
              Yakin ingin menghapus data siswa <strong>{deleteConfirmTarget.name}</strong>? Nilai,
              kehadiran, dan keterkaitan siswa ini akan diperbarui.
            </p>
            <div className="flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setDeleteConfirmTarget(null)}
                className="px-3 py-1.5 text-xs font-semibold border border-slate-300 rounded-lg hover:bg-slate-50 cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={() => {
                  deleteSiswa(deleteConfirmTarget.id);
                  setDeleteConfirmTarget(null);
                }}
                className="px-3 py-1.5 text-xs font-bold bg-red-600 hover:bg-red-700 text-white rounded-lg cursor-pointer"
              >
                Ya, Hapus
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Detail Form Siswa (Tambah / Edit Lengkap) */}
      {modalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-3xl w-full overflow-hidden border border-slate-200">
            <div className="bg-gradient-to-r from-blue-800 to-indigo-800 p-4 text-white flex items-center justify-between">
              <h3 className="font-bold text-sm flex items-center gap-2">
                <GraduationCap size={18} />{' '}
                {editingSiswa ? 'Edit Identitas Siswa Lengkap' : 'Tambah Peserta Didik Baru'}
              </h3>
              <button
                type="button"
                onClick={() => setModalOpen(false)}
                className="text-white/80 hover:text-white cursor-pointer"
                aria-label="Tutup modal"
              >
                <X size={18} />
              </button>
            </div>

            <form
              onSubmit={handleModalSubmit}
              className="p-5 space-y-4 max-h-[80vh] overflow-y-auto text-xs"
            >
              <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">NIS</label>
                  <input
                    type="text"
                    value={formData.nis}
                    onChange={(e) => setFormData({ ...formData, nis: e.target.value })}
                    placeholder="26001"
                    className="w-full px-3 py-2 border rounded-lg font-mono"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">NISN</label>
                  <input
                    type="text"
                    value={formData.nisn}
                    onChange={(e) => setFormData({ ...formData, nisn: e.target.value })}
                    placeholder="Opsional"
                    className="w-full px-3 py-2 border rounded-lg font-mono"
                  />
                </div>
                <div className="sm:col-span-2">
                  <label className="block font-semibold text-slate-700 mb-1">
                    Nama Lengkap Siswa *
                  </label>
                  <input
                    type="text"
                    required
                    value={formData.namaSiswa}
                    onChange={(e) => setFormData({ ...formData, namaSiswa: e.target.value })}
                    placeholder="Nama Lengkap"
                    className="w-full px-3 py-2 border rounded-lg font-bold"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Panggilan</label>
                  <input
                    type="text"
                    value={formData.namaPanggilan || ''}
                    onChange={(e) => setFormData({ ...formData, namaPanggilan: e.target.value })}
                    placeholder="Nama Panggilan"
                    className="w-full px-3 py-2 border rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Tempat Lahir</label>
                  <input
                    type="text"
                    value={formData.tempatLahir}
                    onChange={(e) => setFormData({ ...formData, tempatLahir: e.target.value })}
                    placeholder="Oksibil"
                    className="w-full px-3 py-2 border rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Tanggal Lahir</label>
                  <input
                    type="date"
                    value={formData.tanggalLahir}
                    onChange={(e) => setFormData({ ...formData, tanggalLahir: e.target.value })}
                    className="w-full px-3 py-2 border rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Jenis Kelamin</label>
                  <select
                    value={formData.jenisKelamin}
                    onChange={(e) =>
                      setFormData({ ...formData, jenisKelamin: e.target.value as 'L' | 'P' })
                    }
                    className="w-full px-3 py-2 border rounded-lg"
                  >
                    <option value="L">Laki-laki (L)</option>
                    <option value="P">Perempuan (P)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Agama</label>
                  <select
                    value={formData.agama}
                    onChange={(e) => setFormData({ ...formData, agama: e.target.value })}
                    className="w-full px-3 py-2 border rounded-lg bg-cyan-50 font-bold"
                  >
                    {AGAMA_OPTIONS.map((ag) => (
                      <option key={ag} value={ag}>
                        {ag}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">
                    Pendidikan Sebelumnya
                  </label>
                  <input
                    type="text"
                    value={formData.sekolahAsal || ''}
                    onChange={(e) => setFormData({ ...formData, sekolahAsal: e.target.value })}
                    placeholder="TK / PAUD / Asal Sekolah"
                    className="w-full px-3 py-2 border rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Kelas</label>
                  <select
                    value={formData.kelas}
                    onChange={(e) => {
                      const k = e.target.value as ClassLevel;
                      setFormData({ ...formData, kelas: k, fase: getFase(k) });
                    }}
                    className="w-full px-3 py-2 border rounded-lg font-bold"
                  >
                    {ALL_CLASSES.map((k) => (
                      <option key={k} value={k}>
                        Kelas {k}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Status</label>
                  <select
                    value={formData.statusSiswa}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        statusSiswa: e.target.value as 'Aktif' | 'Mutasi' | 'Lulus'
                      })
                    }
                    className="w-full px-3 py-2 border rounded-lg"
                  >
                    <option value="Aktif">Aktif</option>
                    <option value="Mutasi">Mutasi</option>
                    <option value="Lulus">Lulus</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-semibold text-slate-700 mb-1">
                  Alamat Peserta Didik
                </label>
                <input
                  type="text"
                  value={formData.alamat}
                  onChange={(e) => setFormData({ ...formData, alamat: e.target.value })}
                  placeholder="Alamat peserta didik"
                  className="w-full px-3 py-2 border rounded-lg"
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Nama Ayah</label>
                  <input
                    type="text"
                    value={formData.namaAyah}
                    onChange={(e) => setFormData({ ...formData, namaAyah: e.target.value })}
                    className="w-full px-3 py-2 border rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Nama Ibu</label>
                  <input
                    type="text"
                    value={formData.namaIbu}
                    onChange={(e) => setFormData({ ...formData, namaIbu: e.target.value })}
                    className="w-full px-3 py-2 border rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Pekerjaan Ayah</label>
                  <input
                    type="text"
                    value={formData.pekerjaanAyah || ''}
                    onChange={(e) => setFormData({ ...formData, pekerjaanAyah: e.target.value })}
                    className="w-full px-3 py-2 border rounded-lg"
                  />
                </div>
                <div>
                  <label className="block font-semibold text-slate-700 mb-1">Pekerjaan Ibu</label>
                  <input
                    type="text"
                    value={formData.pekerjaanIbu || ''}
                    onChange={(e) => setFormData({ ...formData, pekerjaanIbu: e.target.value })}
                    className="w-full px-3 py-2 border rounded-lg"
                  />
                </div>
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                <div className="font-bold text-slate-700 text-[11px] uppercase">
                  Alamat Orang Tua
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <input
                    type="text"
                    value={formData.alamatOrtuJalan || ''}
                    onChange={(e) =>
                      setFormData({ ...formData, alamatOrtuJalan: e.target.value })
                    }
                    placeholder="Jalan"
                    className="px-2.5 py-1.5 border rounded-lg bg-white"
                  />
                  <input
                    type="text"
                    value={formData.alamatOrtuDesa || ''}
                    onChange={(e) =>
                      setFormData({ ...formData, alamatOrtuDesa: e.target.value })
                    }
                    placeholder="Kelurahan/Desa"
                    className="px-2.5 py-1.5 border rounded-lg bg-white"
                  />
                  <input
                    type="text"
                    value={formData.alamatOrtuKecamatan || ''}
                    onChange={(e) =>
                      setFormData({ ...formData, alamatOrtuKecamatan: e.target.value })
                    }
                    placeholder="Kecamatan"
                    className="px-2.5 py-1.5 border rounded-lg bg-white"
                  />
                  <input
                    type="text"
                    value={formData.alamatOrtuKabupaten || ''}
                    onChange={(e) =>
                      setFormData({ ...formData, alamatOrtuKabupaten: e.target.value })
                    }
                    placeholder="Kabupaten/Kota"
                    className="px-2.5 py-1.5 border rounded-lg bg-white"
                  />
                  <input
                    type="text"
                    value={formData.alamatOrtuProvinsi || ''}
                    onChange={(e) =>
                      setFormData({ ...formData, alamatOrtuProvinsi: e.target.value })
                    }
                    placeholder="Provinsi"
                    className="px-2.5 py-1.5 border rounded-lg bg-white sm:col-span-2"
                  />
                </div>
              </div>

              <div className="p-3 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
                <div className="font-bold text-slate-700 text-[11px] uppercase">
                  Wali Peserta Didik
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <input
                    type="text"
                    value={formData.namaWali || ''}
                    onChange={(e) => setFormData({ ...formData, namaWali: e.target.value })}
                    placeholder="Nama Wali"
                    className="px-2.5 py-1.5 border rounded-lg bg-white"
                  />
                  <input
                    type="text"
                    value={formData.pekerjaanWali || ''}
                    onChange={(e) => setFormData({ ...formData, pekerjaanWali: e.target.value })}
                    placeholder="Pekerjaan Wali"
                    className="px-2.5 py-1.5 border rounded-lg bg-white"
                  />
                  <input
                    type="text"
                    value={formData.alamatWali || ''}
                    onChange={(e) => setFormData({ ...formData, alamatWali: e.target.value })}
                    placeholder="Alamat Wali"
                    className="px-2.5 py-1.5 border rounded-lg bg-white"
                  />
                </div>
              </div>

              <div className="pt-3 border-t flex items-center justify-end gap-2">
                <SaveStatusIndicator status={modalSaveStatus} compact />
                <button
                  type="button"
                  onClick={() => setModalOpen(false)}
                  className="px-4 py-2 border border-slate-300 rounded-lg font-medium text-xs hover:bg-slate-50 cursor-pointer"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  disabled={modalSaveStatus === 'saving'}
                  className="px-4 py-2 bg-blue-700 hover:bg-blue-800 text-white rounded-lg font-bold flex items-center gap-1.5 text-xs cursor-pointer disabled:opacity-60"
                >
                  <Save size={14} /> {editingSiswa ? 'Update Siswa' : 'Simpan Siswa'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
