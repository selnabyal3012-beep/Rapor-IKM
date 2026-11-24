import React, { useState, useEffect, useRef } from 'react';
import {
  Home,
  HelpCircle,
  Trash2,
  Save,
  Upload,
  FileSpreadsheet,
  FolderDown,
  Image as ImageIcon,
  CheckCircle2,
  Link2,
  Building2,
  RotateCcw,
  AlertCircle,
  X
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { SchoolIdentity, ClassLevel, FaseType, SemesterType, ALL_CLASSES } from '../../types';
import { useSaveFeedback } from '../../hooks/useSaveFeedback';
import { SaveStatusIndicator } from '../common/SaveStatusIndicator';
import { parseAndImportTemplate } from '../../utils/templateImporter';
import * as XLSX from 'xlsx';

const VALID_FASE: FaseType[] = ['A', 'B', 'C', 'D', 'E', 'F'];

const PointingHandRight: React.FC<{ className?: string }> = ({ className = 'w-6 h-5' }) => (
  <svg viewBox="0 0 36 24" className={className} fill="none" aria-hidden="true">
    <path
      d="M2 10h11l-1-4c-.2-1 .8-2 1.8-1.6l4.2 1.8 2.2 2.8H32c1.4 0 2.5 1 2.5 2.2S33.4 13.5 32 13.5h-9.2c.6.8.7 2 .1 2.8-.6.7-1.4.9-2.2.9.3.8.1 1.8-.5 2.4-.7.6-1.6.7-2.4.5 0 .9-.6 1.7-1.6 1.9H9.5C5.8 22 2 19.5 2 15.2V10z"
      fill="#f8fafc"
      stroke="#334155"
      strokeWidth="1.6"
      strokeLinejoin="round"
    />
  </svg>
);

const PointingHandDownLeft: React.FC<{ className?: string }> = ({ className = 'w-9 h-9' }) => (
  <svg viewBox="0 0 36 36" className={className} fill="none" aria-hidden="true">
    <path
      d="M22 4l6 6-7.5 7.5 3 3-13.5 9.5 3.5-14.5 3 2.5L22 4z"
      fill="#ffffff"
      stroke="#1e293b"
      strokeWidth="1.7"
      strokeLinejoin="round"
    />
  </svg>
);

const buildInitialForm = (sekolah: SchoolIdentity): SchoolIdentity => {
  const combinedTempatTanggal =
    sekolah.tempatTanggalRapor?.trim() ||
    [sekolah.tempatRapor, sekolah.tanggalRapor].filter((v) => v && v.trim() !== '').join(', ');
  return {
    ...sekolah,
    kabupatenLabel: sekolah.kabupatenLabel || 'Kabupaten',
    faks: sekolah.faks ?? '-',
    kelasParalel: sekolah.kelasParalel ?? '',
    tempatTanggalRapor: combinedTempatTanggal
  };
};

export const SchoolIdentitySection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    updateSekolah,
    setSelectedKelas,
    setSelectedFase,
    setSelectedSemester,
    setActiveMenu,
    openClearDataModal,
    showSaveNotification
  } = appContext;

  const { status: saveStatus, triggerSave, lastSavedTime } = useSaveFeedback();
  const [formData, setFormData] = useState<SchoolIdentity>(() => buildInitialForm(db.sekolah));
  const [showHelp, setShowHelp] = useState(false);
  const [importError, setImportError] = useState<string | null>(null);
  const [savedBanner, setSavedBanner] = useState(false);

  const kepalaInputRef = useRef<HTMLInputElement>(null);
  const nipKepalaInputRef = useRef<HTMLInputElement>(null);
  const nipGuruInputRef = useRef<HTMLInputElement>(null);
  const importInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    setFormData(buildInitialForm(db.sekolah));
  }, [db.sekolah]);

  const activeClassCount = db.siswa.filter((s) => s.kelas === (formData.kelas || 'I')).length;
  const baseKelas = String(formData.kelas || 'I').trim();
  const parallelSuffix = formData.kelasParalel ?? '';
  const computedRombel = parallelSuffix.trim() ? `${baseKelas}${parallelSuffix}` : baseKelas;

  const studentCountOptions = Array.from(
    new Set([
      ...(activeClassCount > 0 ? [activeClassCount] : []),
      ...(typeof formData.jumlahSiswaProfil === 'number' && formData.jumlahSiswaProfil > 0
        ? [formData.jumlahSiswaProfil]
        : []),
      ...Array.from({ length: 65 }, (_, idx) => idx + 1)
    ])
  ).sort((a, b) => a - b);

  const handleFieldChange = (name: keyof SchoolIdentity, value: string | number) => {
    setImportError(null);
    setSavedBanner(false);
    setFormData((prev) => {
      const next: SchoolIdentity = {
        ...prev,
        [name]: value
      };

      if (name === 'tempatTanggalRapor' && typeof value === 'string') {
        const trimmed = value.trim();
        if (trimmed.includes(',')) {
          const [tempatPart, ...tglParts] = trimmed.split(',');
          next.tempatRapor = tempatPart.trim();
          next.tanggalRapor = tglParts.join(',').trim();
        } else {
          next.tanggalRapor = trimmed;
        }
      } else if (name === 'tempatRapor' || name === 'tanggalRapor') {
        const tempat = name === 'tempatRapor' ? String(value) : prev.tempatRapor || '';
        const tanggal = name === 'tanggalRapor' ? String(value) : prev.tanggalRapor || '';
        next.tempatTanggalRapor = [tempat, tanggal].filter((v) => v && v.trim() !== '').join(', ');
      }

      if (name === 'kelas' || name === 'kelasParalel') {
        const kls = String(name === 'kelas' ? value : prev.kelas || 'I').trim();
        const par = String(name === 'kelasParalel' ? value : prev.kelasParalel ?? '');
        next.rombel = par.trim() ? `${kls}${par}` : kls;
      }

      return next;
    });
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    if (name === 'jumlahSiswaProfil') {
      const parsed = parseInt(value, 10);
      handleFieldChange('jumlahSiswaProfil', Number.isNaN(parsed) ? 0 : parsed);
      return;
    }
    handleFieldChange(name as keyof SchoolIdentity, value);
  };

  const handleUploadLogo = (field: 'logoKiriUrl' | 'logoKananUrl') => (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const result = event.target?.result;
      if (typeof result === 'string') {
        handleFieldChange(field, result);
      }
    };
    reader.readAsDataURL(file);
  };

  const handleSubmit = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const payload: SchoolIdentity = {
      ...formData,
      rombel: computedRombel
    };
    triggerSave(() => {
      updateSekolah(payload);
      if (payload.kelas && ALL_CLASSES.includes(payload.kelas as ClassLevel)) {
        setSelectedKelas(payload.kelas as ClassLevel);
      }
      if (payload.fase && VALID_FASE.includes(payload.fase as FaseType)) {
        setSelectedFase(payload.fase as FaseType);
      }
      if (payload.semester) {
        setSelectedSemester(payload.semester as SemesterType);
      }
      setSavedBanner(true);
      showSaveNotification(
        `Profil Sekolah disimpan & rumus identitas dihubungkan ke Kelas ${computedRombel}, Fase ${payload.fase || '-'}, dan seluruh rapor.`
      );
      setTimeout(() => setSavedBanner(false), 3500);
    });
  };

  const handleResetChanges = () => {
    setFormData(buildInitialForm(db.sekolah));
    setImportError(null);
    showSaveNotification('Isian dikembalikan ke data Profil Sekolah tersimpan.');
  };

  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const rows: (string | number)[][] = [
      ['PROFIL SEKOLAH & IDENTITAS TERHUBUNG'],
      ['Parameter', 'Nilai / Keterangan'],
      ['Nama Pemerintah', formData.namaPemerintah || ''],
      ['Nama Instansi', formData.namaInstansi || ''],
      ['Nama Sekolah', formData.namaSekolah || ''],
      ['NISN/NIS/NSS', formData.nss || ''],
      ['NPSN', formData.npsn || ''],
      ['Alamat', formData.alamat || ''],
      ['Desa/Kelurahan', formData.desaKelurahan || ''],
      ['Kecamatan', formData.distrikKecamatan || ''],
      [formData.kabupatenLabel || 'Kabupaten', formData.kabupaten || ''],
      ['Provinsi', formData.provinsi || ''],
      ['Faks', formData.faks || ''],
      ['Kode POS', formData.kodePos || ''],
      ['Telp.', formData.telepon || ''],
      ['Tahun Pelajaran', formData.tahunPelajaran || ''],
      ['Semester', formData.semester || ''],
      ['Kepala Sekolah', formData.namaKepalaSekolah || ''],
      ['NIP Kepala Sekolah', formData.nipKepalaSekolah || ''],
      ['Guru Kelas', formData.waliKelas || ''],
      ['NIP Guru Kelas', formData.nipWaliKelas || ''],
      ['Website Sekolah', formData.website || ''],
      ['E-mail Sekolah', formData.emailSekolah || ''],
      ['Tempat/Tanggal Rapor', formData.tempatTanggalRapor || ''],
      ['Kelas', formData.kelas || ''],
      ['Kelas Paralel', formData.kelasParalel || ''],
      ['Rombel Gabungan (Rumus)', computedRombel],
      ['Jumlah Siswa', formData.jumlahSiswaProfil ?? activeClassCount],
      ['Fase', formData.fase || '']
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 28 }, { wch: 48 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Profil Sekolah');
    XLSX.writeFile(wb, `profil-sekolah-${formData.npsn || 'sdn-balil'}.xlsx`);
  };

  const handleImportFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    try {
      const result = await parseAndImportTemplate(file, 'template_sekolah', appContext);
      if (result.success) {
        setImportError(null);
        showSaveNotification(result.message);
      } else {
        setImportError(result.message);
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Format file tidak valid';
      setImportError(`Gagal membaca file: ${msg}`);
    }
  };

  const semesterDisplay =
    formData.semester === 'Genap' || formData.semester === '2' ? 'Genap' : 'Ganjil';

  return (
    <div
      className="min-h-[calc(100vh-5rem)] rounded-2xl border-2 border-red-600 p-3 sm:p-5 shadow-md"
      style={{
        backgroundColor: '#00ff00',
        backgroundImage: 'radial-gradient(#065f46 0.85px, transparent 0.85px)',
        backgroundSize: '8px 8px'
      }}
    >
      {/* Top Action Strip + Black Corner Icon Bar matching image_3ae5448e.png */}
      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div className="flex flex-wrap items-center gap-2 bg-white/95 border-2 border-red-600 rounded-xl px-3 py-2 shadow-sm">
          <button
            type="button"
            onClick={() => handleSubmit()}
            disabled={saveStatus === 'saving'}
            className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold rounded-lg shadow-xs transition-colors cursor-pointer disabled:opacity-60"
          >
            <Save size={14} />
            <span>Simpan &amp; Hubungkan Rumus</span>
          </button>

          <button
            type="button"
            onClick={handleResetChanges}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 text-xs font-semibold rounded-lg transition-colors cursor-pointer"
          >
            <RotateCcw size={13} />
            <span>Reset</span>
          </button>

          <button
            type="button"
            onClick={() => importInputRef.current?.click()}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200 text-xs font-semibold rounded-lg transition-colors cursor-pointer"
          >
            <Upload size={13} />
            <span>Impor Excel</span>
          </button>
          <input
            ref={importInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleImportFile}
            className="hidden"
          />

          <button
            type="button"
            onClick={handleExportExcel}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-900 border border-emerald-300 text-xs font-semibold rounded-lg transition-colors cursor-pointer"
          >
            <FileSpreadsheet size={13} />
            <span>Expor .xlsx</span>
          </button>

          <button
            type="button"
            onClick={() => setActiveMenu('template_data')}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 text-xs font-semibold rounded-lg transition-colors cursor-pointer"
            title="Unduh Template Excel Profil Sekolah"
          >
            <FolderDown size={13} />
            <span>Template</span>
          </button>

          <SaveStatusIndicator status={saveStatus} compact lastSavedText={lastSavedTime} />

          {savedBanner && (
            <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md bg-emerald-900 text-white text-[11px] font-bold">
              <CheckCircle2 size={13} className="text-emerald-300" />
              Rumus Identitas Tersinkron
            </span>
          )}
        </div>

        {/* Black Corner Bar: Home, ?, Trash */}
        <div className="bg-black border-2 border-red-600 px-3 py-1.5 flex items-center gap-3 shadow-md ml-auto">
          <button
            type="button"
            onClick={() => setActiveMenu('dashboard')}
            title="Kembali ke Beranda / Dashboard"
            aria-label="Beranda"
            className="w-9 h-9 rounded-full bg-gradient-to-b from-sky-400 to-blue-700 hover:brightness-110 border-2 border-sky-200 text-white flex items-center justify-center shadow-xs transition-transform active:scale-95 cursor-pointer"
          >
            <Home size={18} />
          </button>

          <button
            type="button"
            onClick={() => setShowHelp((prev) => !prev)}
            title="Panduan Rumus Identitas Sekolah"
            aria-label="Panduan Identitas Sekolah"
            className="w-9 h-9 rounded-full bg-gradient-to-b from-purple-400 to-purple-800 hover:brightness-110 border-2 border-purple-200 text-white flex items-center justify-center shadow-xs transition-transform active:scale-95 cursor-pointer"
          >
            <HelpCircle size={20} />
          </button>

          <button
            type="button"
            onClick={openClearDataModal}
            title="Hapus / Kosongkan Data Aplikasi"
            aria-label="Hapus Data"
            className="w-9 h-9 rounded-full bg-gradient-to-b from-slate-100 to-slate-300 hover:bg-rose-50 border-2 border-red-600 text-red-700 flex items-center justify-center shadow-xs transition-transform active:scale-95 cursor-pointer"
          >
            <Trash2 size={17} />
          </button>
        </div>
      </div>

      {/* Inline Error Banner (if Excel import fails) */}
      {importError && (
        <div className="mb-4 bg-rose-50 border-2 border-rose-600 text-rose-900 rounded-xl px-4 py-2.5 flex items-center justify-between gap-3 text-xs font-semibold shadow-xs">
          <div className="flex items-center gap-2">
            <AlertCircle size={16} className="text-rose-600 shrink-0" />
            <span>{importError}</span>
          </div>
          <button
            type="button"
            onClick={() => setImportError(null)}
            className="text-rose-700 hover:text-rose-950 p-1 cursor-pointer"
            aria-label="Tutup pesan error"
          >
            <X size={14} />
          </button>
        </div>
      )}

      {/* Help Drawer */}
      {showHelp && (
        <div className="mb-4 bg-white border-2 border-purple-700 rounded-xl p-4 shadow-md text-xs text-slate-800 space-y-2">
          <div className="flex items-center justify-between border-b border-slate-200 pb-2">
            <h3 className="font-bold text-sm text-purple-900 flex items-center gap-2">
              <Link2 size={16} /> Cara Kerja Rumus Otomatis PROFIL SEKOLAH
            </h3>
            <button
              type="button"
              onClick={() => setShowHelp(false)}
              className="text-slate-500 hover:text-slate-900 p-1 cursor-pointer"
              aria-label="Tutup panduan"
            >
              <X size={15} />
            </button>
          </div>
          <ul className="list-disc pl-5 space-y-1 leading-relaxed">
            <li>
              Seluruh isian pada tabel <strong>PROFIL SEKOLAH</strong> (Nama Sekolah, NPSN, Alamat, Kepala Sekolah, Guru Kelas &amp; NIP, Tahun Pelajaran, Semester, Tempat/Tanggal Rapor, Kelas/Rombel, dan Fase) otomatis dihubungkan ke <strong>Rapor Akademik, Rapor PTS, Rapor P5, Buku Induk, Piagam, KKM, Data Siswa, Sidebar, dan Dashboard</strong>.
            </li>
            <li>
              <strong>Rumus Kelas Paralel:</strong> Pilih tingkat kelas pada baris <em>Kelas</em> (misalnya <code>I</code>), lalu isi akhiran paralel pada kotak merah di sebelah kanan (misalnya <code>B</code>, <code>.B</code>, atau <code>/B</code>). Hasil gabungan seperti <code>IB</code> otomatis muncul di sebelah pilihan Kelas dan tersimpan sebagai Rombel.
            </li>
            <li>
              <strong>Ganti Label Wilayah &amp; Pimpinan:</strong> Klik dropdown <em>Kabupaten / Kota</em> untuk mengganti sebutan wilayah, serta isi langsung nama Kepala Sekolah, Guru Kelas, dan NIP tanpa perlu menu Data Guru terpisah.
            </li>
          </ul>
        </div>
      )}

      <form onSubmit={handleSubmit}>
        <div className="grid grid-cols-1 xl:grid-cols-12 gap-5 items-start">
          {/* LEFT / MAIN EXCEL PROFIL SEKOLAH TABLE */}
          <div className="xl:col-span-8 bg-white border-2 border-red-600 shadow-lg overflow-hidden">
            {/* Black Header */}
            <div className="bg-black text-white text-center py-2 px-4 border-b-2 border-red-600">
              <h1 className="font-serif font-bold text-2xl sm:text-3xl tracking-wider uppercase">
                PROFIL SEKOLAH
              </h1>
            </div>

            <div className="divide-y divide-red-600 text-[13px] sm:text-sm font-serif">
              {/* Row 1: Nama Instansi */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Nama Instansi
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="namaInstansi"
                  value={formData.namaInstansi || ''}
                  onChange={handleInputChange}
                  placeholder="Dinas Pendidikan, Kearsipan Dan Perpustakaan Daerah"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 2: Nama Sekolah */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Nama Sekolah
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="namaSekolah"
                  required
                  value={formData.namaSekolah || ''}
                  onChange={handleInputChange}
                  placeholder="SD Negeri Balil"
                  className="flex-1 min-w-0 bg-white px-2 py-1 font-bold text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 3: NISN/NIS/NSS */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  NISN/NIS/NSS
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="nss"
                  value={formData.nss || ''}
                  onChange={handleInputChange}
                  placeholder="_"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 4: NPSN */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  NPSN
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="npsn"
                  value={formData.npsn || ''}
                  onChange={handleInputChange}
                  placeholder="_"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 5: Alamat */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Alamat
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="alamat"
                  value={formData.alamat || ''}
                  onChange={handleInputChange}
                  placeholder="Jalan Betaabib Kabiding"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 6: Desa/Kelurahan */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Desa/Kelurahan
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="desaKelurahan"
                  value={formData.desaKelurahan || ''}
                  onChange={handleInputChange}
                  placeholder="Kabiding"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 7: Kecamatan */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Kecamatan
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="distrikKecamatan"
                  value={formData.distrikKecamatan || ''}
                  onChange={handleInputChange}
                  placeholder="Oksibil"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 8: Ganti Kabupaten / Kota */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-white px-2 py-1 border-r border-red-600 flex items-center justify-between gap-1">
                  <span className="inline-flex items-center gap-0.5 text-[11px] font-sans font-extrabold text-red-600 shrink-0">
                    Ganti
                    <PointingHandRight className="w-5 h-4" />
                  </span>
                  <select
                    name="kabupatenLabel"
                    value={formData.kabupatenLabel || 'Kabupaten'}
                    onChange={handleInputChange}
                    aria-label="Jenis wilayah Kabupaten atau Kota"
                    className="flex-1 min-w-0 border border-slate-400 bg-white px-1.5 py-0.5 text-xs font-sans font-semibold text-black focus:outline-hidden cursor-pointer"
                  >
                    <option value="Kabupaten">Kabupaten</option>
                    <option value="Kota">Kota</option>
                    <option value="Kabupaten/Kota">Kabupaten/Kota</option>
                  </select>
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="kabupaten"
                  value={formData.kabupaten || ''}
                  onChange={handleInputChange}
                  placeholder="Pegunungan Bintang"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 9: Provinsi */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Provinsi
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="provinsi"
                  value={formData.provinsi || ''}
                  onChange={handleInputChange}
                  placeholder="Papua Pegunungan"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 10: Faks */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Faks
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="faks"
                  value={formData.faks ?? ''}
                  onChange={handleInputChange}
                  placeholder="_"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 11: Kode POS */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Kode POS
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="kodePos"
                  value={formData.kodePos || ''}
                  onChange={handleInputChange}
                  placeholder="_"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 12: Telp. */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Telp.
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="telepon"
                  value={formData.telepon || ''}
                  onChange={handleInputChange}
                  placeholder="_"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 13: Tahun Pelajaran */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Tahun Pelajaran
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <div className="flex-1 min-w-0 bg-white flex items-center">
                  <input
                    type="text"
                    name="tahunPelajaran"
                    list="profil-tahun-pelajaran-list"
                    value={formData.tahunPelajaran || ''}
                    onChange={handleInputChange}
                    placeholder="2026/2027"
                    className="w-full bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                  />
                  <datalist id="profil-tahun-pelajaran-list">
                    <option value="2025/2026" />
                    <option value="2026" />
                    <option value="2026/2027" />
                    <option value="2027/2028" />
                  </datalist>
                </div>
              </div>

              {/* Row 14: Semester */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Semester
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <div className="flex-1 min-w-0 bg-[#e5e5e5] px-1 py-0.5 flex items-center">
                  <select
                    name="semester"
                    value={semesterDisplay}
                    onChange={handleInputChange}
                    className="w-28 sm:w-32 border border-slate-500 bg-white px-2 py-0.5 font-sans text-xs font-bold text-black focus:outline-hidden cursor-pointer"
                  >
                    <option value="Ganjil">I (Ganjil)</option>
                    <option value="Genap">II (Genap)</option>
                  </select>
                </div>
              </div>

              {/* Row 15: Ganti Kepala Sekolah */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-white px-2 py-1.5 border-r border-red-600 flex items-center justify-between gap-2">
                  <button
                    type="button"
                    onClick={() => kepalaInputRef.current?.focus()}
                    className="inline-flex items-center gap-0.5 text-[11px] font-sans font-extrabold text-red-600 hover:underline cursor-pointer shrink-0"
                  >
                    Ganti
                    <PointingHandRight className="w-5 h-4" />
                  </button>
                  <span className="font-bold italic text-black flex-1">Kepala Sekolah</span>
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  ref={kepalaInputRef}
                  type="text"
                  name="namaKepalaSekolah"
                  required
                  value={formData.namaKepalaSekolah || ''}
                  onChange={handleInputChange}
                  placeholder="Wihelmina Rumbiak, S.Pd.Gr"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 16: Ganti NIP Kepala Sekolah */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-white border-r border-red-600 flex items-stretch">
                  <button
                    type="button"
                    onClick={() => nipKepalaInputRef.current?.focus()}
                    className="w-[92px] sm:w-[104px] shrink-0 bg-[#76f5b0] px-2 py-1.5 flex items-center gap-0.5 text-[11px] font-sans font-extrabold text-red-600 hover:brightness-95 cursor-pointer"
                  >
                    <span>Ganti</span>
                    <PointingHandRight className="w-5 h-4" />
                  </button>
                  <div className="flex-1 bg-white px-2 py-1.5 font-bold italic text-black flex items-center">
                    NIP.
                  </div>
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  ref={nipKepalaInputRef}
                  type="text"
                  name="nipKepalaSekolah"
                  value={formData.nipKepalaSekolah || ''}
                  onChange={handleInputChange}
                  placeholder="-"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 17: Guru Kelas */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Guru Kelas
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="waliKelas"
                  required
                  value={formData.waliKelas || ''}
                  onChange={handleInputChange}
                  placeholder="Selpius Nabyal, S.Pd.Gr"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 18: Ganti NIP Guru Kelas */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-white border-r border-red-600 flex items-stretch">
                  <button
                    type="button"
                    onClick={() => nipGuruInputRef.current?.focus()}
                    className="w-[92px] sm:w-[104px] shrink-0 bg-[#76f5b0] px-2 py-1.5 flex items-center gap-0.5 text-[11px] font-sans font-extrabold text-red-600 hover:brightness-95 cursor-pointer"
                  >
                    <span>Ganti</span>
                    <PointingHandRight className="w-5 h-4" />
                  </button>
                  <div className="flex-1 bg-white px-2 py-1.5 font-bold italic text-black flex items-center">
                    NIP.
                  </div>
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  ref={nipGuruInputRef}
                  type="text"
                  name="nipWaliKelas"
                  value={formData.nipWaliKelas || ''}
                  onChange={handleInputChange}
                  placeholder="199312302024211027"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 19: Website Sekolah */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Website Sekolah
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="website"
                  value={formData.website || ''}
                  onChange={handleInputChange}
                  placeholder="_"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-blue-700 underline decoration-blue-700/60 focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 20: E-mail Sekolah */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  E-mail Sekolah
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="emailSekolah"
                  value={formData.emailSekolah || ''}
                  onChange={handleInputChange}
                  placeholder="_"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-blue-700 underline decoration-blue-700/60 focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 21: Tempat/Tanggal Rapor */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Tempat/Tanggal Rapor
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <input
                  type="text"
                  name="tempatTanggalRapor"
                  value={formData.tempatTanggalRapor || ''}
                  onChange={handleInputChange}
                  placeholder="Oksibil, 10 Juli 2026"
                  className="flex-1 min-w-0 bg-white px-2 py-1 text-black focus:bg-yellow-50 focus:outline-hidden"
                />
              </div>

              {/* Row 22: Kelas + Live Rumus Kelas Paralel */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Kelas
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <div className="flex-1 min-w-0 bg-[#e5e5e5] px-1 py-0.5 flex items-center gap-3">
                  <select
                    name="kelas"
                    value={formData.kelas || 'I'}
                    onChange={handleInputChange}
                    aria-label="Pilih tingkat kelas"
                    className="w-24 border border-slate-500 bg-white px-2 py-0.5 font-sans text-xs font-bold text-black focus:outline-hidden cursor-pointer"
                  >
                    {ALL_CLASSES.map((kls) => (
                      <option key={kls} value={kls}>
                        {kls}
                      </option>
                    ))}
                  </select>
                  <div
                    className="flex-1 text-center font-serif font-bold text-base text-black tracking-wide select-none"
                    title="Rumus otomatis Kelas + Kelas Paralel"
                  >
                    {computedRombel}
                  </div>
                </div>
              </div>

              {/* Row 23: Jumlah Siswa */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#00ffff] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Jumlah Siswa
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <div className="flex-1 min-w-0 bg-[#e5e5e5] px-1 py-0.5 flex flex-wrap items-center justify-between gap-2">
                  <select
                    name="jumlahSiswaProfil"
                    value={formData.jumlahSiswaProfil ?? activeClassCount}
                    onChange={handleInputChange}
                    aria-label="Pilih jumlah siswa"
                    className="w-24 border border-slate-500 bg-white px-2 py-0.5 font-sans text-xs font-bold text-black focus:outline-hidden cursor-pointer"
                  >
                    {studentCountOptions.map((n) => (
                      <option key={n} value={n}>
                        {n}
                      </option>
                    ))}
                  </select>

                  <button
                    type="button"
                    onClick={() => handleFieldChange('jumlahSiswaProfil', activeClassCount)}
                    className="px-2 py-0.5 bg-white hover:bg-cyan-50 border border-slate-400 rounded text-[11px] font-sans font-semibold text-slate-800 cursor-pointer"
                    title="Cocokkan dengan jumlah siswa terdaftar di kelas ini"
                  >
                    Sesuaikan Siswa Kelas {baseKelas} ({activeClassCount})
                  </button>
                </div>
              </div>

              {/* Row 24: Fase */}
              <div className="flex items-stretch">
                <div className="w-[170px] sm:w-[235px] shrink-0 bg-[#76f5b0] px-2.5 py-1.5 font-bold italic text-black border-r border-red-600 flex items-center">
                  Fase
                </div>
                <div className="w-5 shrink-0 bg-white flex items-center justify-center font-bold text-black">
                  :
                </div>
                <div className="flex-1 min-w-0 bg-[#e5e5e5] px-1 py-0.5 flex items-center gap-3">
                  <select
                    name="fase"
                    value={formData.fase || 'A'}
                    onChange={handleInputChange}
                    aria-label="Pilih fase"
                    className="w-24 border border-slate-500 bg-white px-2 py-0.5 font-sans text-xs font-bold text-black focus:outline-hidden cursor-pointer"
                  >
                    <option value="Fondasi">Fondasi</option>
                    <option value="A">A</option>
                    <option value="B">B</option>
                    <option value="C">C</option>
                    <option value="D">D</option>
                    <option value="E">E</option>
                    <option value="F">F</option>
                  </select>
                  <span className="text-[11px] font-sans text-slate-700 font-medium">
                    Fase {formData.fase || 'A'} • Terhubung ke TP, KKM &amp; Rapor
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* RIGHT COLUMN: PARALLEL CLASS BOX + KOP/LOGO + LIVE FORMULA CONNECTIONS */}
          <div className="xl:col-span-4 space-y-4">
            {/* Parallel Class Box matching image_3ae5448e.png */}
            <div className="bg-white/95 border-2 border-red-600 p-3 shadow-md space-y-2.5">
              <div className="bg-white border border-slate-300 p-2.5 text-[11px] leading-snug text-black font-sans">
                <p className="font-bold mb-1">
                  Untuk Kelas paralel Ketik Grade kelas (tanpa menulis kelas) pada cell di bawah ini.
                </p>
                <p className="font-bold">Contoh :</p>
                <p>
                  * A, B, C (diwali dengan <span className="text-red-600 font-bold">Spasi /Tidak</span>)
                </p>
                <p>
                  * .A, .B, .C (diawali dengan <span className="text-red-600 font-bold">tanda titik</span>)
                </p>
                <p>
                  * /A, /B, /C ) (diawali dengan <span className="text-red-600 font-bold">garing</span>) dll.
                  sesuai dengan penulisan di satuan pendidikan masing-masing
                </p>
              </div>

              <div className="flex flex-col items-center pt-1">
                <PointingHandDownLeft className="w-9 h-9 -mb-1" />
                <input
                  type="text"
                  name="kelasParalel"
                  value={formData.kelasParalel ?? ''}
                  onChange={handleInputChange}
                  placeholder="B"
                  aria-label="Input Grade Kelas Paralel"
                  className="w-40 border-2 border-red-600 bg-white px-3 py-1.5 text-center font-serif font-bold text-xl text-black focus:bg-yellow-50 focus:outline-hidden shadow-xs"
                />
                <span className="mt-1.5 text-[11px] font-sans font-semibold text-slate-700">
                  Rumus Hasil Kelas: <strong className="text-emerald-800">{computedRombel}</strong>
                </span>
              </div>
            </div>

            {/* KOP Resmi & Logo Sekolah (tetap tersimpan untuk Rapor & Piagam) */}
            <div className="bg-white border-2 border-red-600 p-3.5 shadow-md space-y-3">
              <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                <h2 className="text-xs font-bold uppercase tracking-wide text-slate-900 flex items-center gap-1.5">
                  <Building2 size={14} className="text-emerald-700" />
                  KOP Resmi &amp; Logo Rapor
                </h2>
                <span className="text-[10px] font-semibold text-emerald-800 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">
                  Terhubung ke KOP Cetak
                </span>
              </div>

              <div className="space-y-2.5 text-xs">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Nama Pemerintah (Baris 1 KOP)
                  </label>
                  <input
                    type="text"
                    name="namaPemerintah"
                    value={formData.namaPemerintah || ''}
                    onChange={handleInputChange}
                    placeholder="PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG"
                    className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg font-bold uppercase text-[11px] focus:ring-2 focus:ring-emerald-500 focus:outline-hidden"
                  />
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">Tempat Rapor</label>
                    <input
                      type="text"
                      name="tempatRapor"
                      value={formData.tempatRapor || ''}
                      onChange={handleInputChange}
                      placeholder="Oksibil"
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs focus:ring-2 focus:ring-emerald-500 focus:outline-hidden"
                    />
                  </div>
                  <div>
                    <label className="block font-bold text-slate-700 mb-1">Tanggal Rapor</label>
                    <input
                      type="text"
                      name="tanggalRapor"
                      value={formData.tanggalRapor || ''}
                      onChange={handleInputChange}
                      placeholder="10 Juli 2026"
                      className="w-full px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs focus:ring-2 focus:ring-emerald-500 focus:outline-hidden"
                    />
                  </div>
                </div>

                {/* Upload Logo Kiri & Logo Kanan */}
                <div className="grid grid-cols-2 gap-2.5 pt-1">
                  <div className="p-2 bg-slate-50 border border-slate-200 rounded-lg space-y-1.5 text-center">
                    <span className="block text-[10px] font-bold text-slate-700">Logo Kiri (Pemda)</span>
                    <div className="w-12 h-12 mx-auto rounded-lg border border-dashed border-slate-300 bg-white flex items-center justify-center overflow-hidden">
                      {formData.logoKiriUrl ? (
                        <img src={formData.logoKiriUrl} alt="Logo Kiri" className="w-full h-full object-contain p-0.5" />
                      ) : (
                        <ImageIcon size={18} className="text-slate-400" />
                      )}
                    </div>
                    <div className="flex items-center justify-center gap-1">
                      <label className="px-2 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded text-[10px] font-bold cursor-pointer">
                        Upload
                        <input
                          type="file"
                          accept="image/*"
                          onChange={handleUploadLogo('logoKiriUrl')}
                          className="hidden"
                        />
                      </label>
                      {formData.logoKiriUrl && (
                        <button
                          type="button"
                          onClick={() => handleFieldChange('logoKiriUrl', '')}
                          className="px-1.5 py-1 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded text-[10px] font-bold cursor-pointer"
                        >
                          Hapus
                        </button>
                      )}
                    </div>
                  </div>

                  <div className="p-2 bg-slate-50 border border-slate-200 rounded-lg space-y-1.5 text-center">
                    <span className="block text-[10px] font-bold text-slate-700">Logo Kanan (Sekolah)</span>
                    <div className="w-12 h-12 mx-auto rounded-lg border border-dashed border-slate-300 bg-white flex items-center justify-center overflow-hidden">
                      {formData.logoKananUrl ? (
                        <img src={formData.logoKananUrl} alt="Logo Kanan" className="w-full h-full object-contain p-0.5" />
                      ) : (
                        <ImageIcon size={18} className="text-slate-400" />
                      )}
                    </div>
                    <div className="flex items-center justify-center gap-1">
                      <label className="px-2 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[10px] font-bold cursor-pointer">
                        Upload
                        <input
                          type="file"
                          accept="image/*"
                          onChange={handleUploadLogo('logoKananUrl')}
                          className="hidden"
                        />
                      </label>
                      {formData.logoKananUrl && (
                        <button
                          type="button"
                          onClick={() => handleFieldChange('logoKananUrl', '')}
                          className="px-1.5 py-1 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded text-[10px] font-bold cursor-pointer"
                        >
                          Hapus
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                {/* Live Mini KOP Preview */}
                <div className="p-2.5 rounded-lg bg-slate-50 border border-slate-300 flex items-center gap-2 text-center">
                  <div className="w-9 h-9 shrink-0 flex items-center justify-center">
                    {formData.logoKiriUrl ? (
                      <img src={formData.logoKiriUrl} alt="Kiri" className="max-h-9 max-w-9 object-contain" />
                    ) : (
                      <div className="w-8 h-8 rounded-full border border-slate-300 bg-white text-[8px] flex items-center justify-center text-slate-500">
                        Kiri
                      </div>
                    )}
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="text-[9px] font-bold uppercase text-slate-700 truncate">
                      {formData.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'}
                    </div>
                    <div className="text-[9px] font-bold uppercase text-slate-700 truncate">
                      {formData.namaInstansi || 'DINAS PENDIDIKAN'}
                    </div>
                    <div className="text-[11px] font-black uppercase text-slate-900 truncate">
                      {formData.namaSekolah || 'SD NEGERI BALIL'}
                    </div>
                    <div className="text-[8px] text-slate-600 truncate">
                      {formData.alamat}, {formData.distrikKecamatan}, {formData.kabupatenLabel || 'Kabupaten'} {formData.kabupaten}
                    </div>
                  </div>
                  <div className="w-9 h-9 shrink-0 flex items-center justify-center">
                    {formData.logoKananUrl ? (
                      <img src={formData.logoKananUrl} alt="Kanan" className="max-h-9 max-w-9 object-contain" />
                    ) : (
                      <div className="w-8 h-8 rounded-full border border-slate-300 bg-white text-[8px] flex items-center justify-center text-slate-500">
                        Kanan
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>

            {/* Live Connected Formulas Summary */}
            <div className="bg-white border-2 border-red-600 p-3.5 shadow-md space-y-2">
              <h3 className="text-xs font-bold uppercase tracking-wide text-slate-900 flex items-center gap-1.5 border-b border-slate-200 pb-1.5">
                <Link2 size={14} className="text-indigo-700" />
                Rumus Identitas Terhubung Otomatis
              </h3>
              <div className="grid grid-cols-2 gap-1.5 text-[11px]">
                <div className="bg-slate-50 border border-slate-200 rounded px-2 py-1">
                  <span className="text-slate-500 block text-[10px]">Rapor Akademik &amp; PTS</span>
                  <strong className="text-slate-900 truncate block">{formData.namaSekolah || '-'}</strong>
                </div>
                <div className="bg-slate-50 border border-slate-200 rounded px-2 py-1">
                  <span className="text-slate-500 block text-[10px]">Kelas / Rombel</span>
                  <strong className="text-slate-900 truncate block">
                    {computedRombel} (Fase {formData.fase || 'A'})
                  </strong>
                </div>
                <div className="bg-slate-50 border border-slate-200 rounded px-2 py-1">
                  <span className="text-slate-500 block text-[10px]">Kepala Sekolah</span>
                  <strong className="text-slate-900 truncate block">{formData.namaKepalaSekolah || '-'}</strong>
                </div>
                <div className="bg-slate-50 border border-slate-200 rounded px-2 py-1">
                  <span className="text-slate-500 block text-[10px]">Guru Kelas (TTD Rapor)</span>
                  <strong className="text-slate-900 truncate block">{formData.waliKelas || '-'}</strong>
                </div>
                <div className="bg-slate-50 border border-slate-200 rounded px-2 py-1 col-span-2">
                  <span className="text-slate-500 block text-[10px]">Titimangsa (Rapor, Buku Induk &amp; Piagam)</span>
                  <strong className="text-slate-900 truncate block">
                    {formData.tempatTanggalRapor || [formData.tempatRapor, formData.tanggalRapor].filter(Boolean).join(', ') || '-'}
                  </strong>
                </div>
              </div>
            </div>
          </div>
        </div>
      </form>
    </div>
  );
};
