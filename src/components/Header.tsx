import React, { useState } from 'react';
import {
  GraduationCap,
  Layers,
  Calendar,
  Laptop,
  Power,
  LogIn,
  LogOut,
  UserCheck,
  Award,
  Sparkles,
  BookOpenCheck,
  Wifi,
  WifiOff,
  Download,
  Share2,
  BookOpen,
  Monitor,
  Smartphone,
  Maximize2,
  Minimize2,
  Menu,
  Save,
  Zap,
  RefreshCw,
  Eye,
  SlidersHorizontal,
  EyeOff,
  Cloud,
  CloudUpload,
  Sun,
  Moon,
  Type,
  Settings,
  FileSpreadsheet,
  Users,
  X,
  Trash2
} from 'lucide-react';
import { useApp } from '../context/AppContext';
import { usePWA } from '../hooks/usePWA';
import { ClassLevel, SemesterType } from '../types';
import { DeveloperModal } from './modals/DeveloperModal';
import { InstallPcModal } from './modals/InstallPcModal';
import { CloseAppModal } from './modals/CloseAppModal';
import { LoginRegisterModal } from './modals/LoginRegisterModal';
import { ShareModal } from './modals/ShareModal';
import { SaveStatusIndicator } from './common/SaveStatusIndicator';
import { PWAInstallButton } from './common/PWAInstallControl';

export const Header: React.FC = () => {
  const { isOnline, isOfflineReady } = usePWA();
  const {
    db,
    selectedKelas,
    setSelectedKelas,
    selectedSemester,
    setSelectedSemester,
    selectedFase,
    setSelectedFase,
    selectedStudentFilter,
    setSelectedStudentFilter,
    currentUser,
    logout,
    setActiveMenu,
    displayMode,
    setDisplayMode,
    isFullscreen,
    toggleFullscreen,
    isMobileDrawerOpen,
    setIsMobileDrawerOpen,
    syncAllData,
    saveAllData,
    globalSaveStatus,
    lastSavedTime,
    openReportPreview,
    openClearDataModal,
    actionToolbarMode,
    setActionToolbarMode,
    firebaseUser,
    cloudSyncStatus,
    syncToCloudFirestore,
    lastCloudSyncTime,
    themeMode,
    setThemeMode,
    uiFontSize,
    setUiFontSize,
    uiFontScalePercent
  } = useApp();

  const [showDevModal, setShowDevModal] = useState(false);
  const [showInstallModal, setShowInstallModal] = useState(false);
  const [showShareModal, setShowShareModal] = useState(false);
  const [showCloseModal, setShowCloseModal] = useState(false);
  const [showAuthModal, setShowAuthModal] = useState(false);

  const kelasOptions: ClassLevel[] = ['I', 'II', 'III', 'IV', 'V', 'VI'];

  return (
    <>
      <header className="bg-slate-900 border-b border-slate-800 text-white shadow-lg sticky top-0 z-40">
        {/* Top bar with Branding & Developer info */}
        <div className="px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 border-b border-slate-800/80 bg-slate-950/70">
          <div className="flex items-center gap-3">
            {/* Mobile Hamburger Drawer Button */}
            <button
              onClick={() => setIsMobileDrawerOpen((prev) => !prev)}
              className={`p-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors cursor-pointer ${
                displayMode === 'hp' ? 'flex' : 'flex lg:hidden'
              } items-center justify-center shrink-0`}
              title="Buka Menu Navigasi"
            >
              <Menu size={18} />
            </button>

            {/* School Crest / Logo Badge */}
            <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-amber-500 via-emerald-600 to-teal-500 p-0.5 shadow-md shrink-0 flex items-center justify-center">
              <div className="w-full h-full bg-slate-900 rounded-[10px] flex items-center justify-center">
                <GraduationCap className="text-amber-400" size={22} />
              </div>
            </div>

            <div>
              <div className="flex items-center gap-2">
                <h1 className="font-extrabold text-sm sm:text-base tracking-wide text-white flex items-center gap-1.5">
                  <span>{db.sekolah.namaSekolah}</span>
                  <span className="text-[10px] bg-emerald-500/20 text-emerald-300 font-semibold px-2 py-0.5 rounded-full border border-emerald-500/30">
                    TP {db.sekolah.tahunPelajaran}
                  </span>
                </h1>
              </div>
              <p className="text-[11px] text-slate-400 hidden sm:block">
                Aplikasi Rapor Kurikulum Merdeka • Distrik Oksibil, Kab. Pegunungan Bintang, Papua Pegunungan
              </p>
            </div>
          </div>

          <div className="flex items-center gap-1.5 flex-wrap">
            {/* PENGATURAN TAMPILAN: PC, HP, FULL */}
            <div className="flex items-center bg-slate-800/90 p-0.5 rounded-lg border border-slate-700/80 shadow-xs">
              <button
                onClick={() => setDisplayMode('pc')}
                title="Tampilan Mode PC (Desktop): Tampilan standar monitor & laptop guru"
                aria-label="Mode Tampilan PC"
                className={`w-7 h-7 inline-flex items-center justify-center rounded-md transition-all cursor-pointer ${
                  displayMode === 'pc'
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white hover:bg-slate-700/60'
                }`}
              >
                <Monitor size={14} />
              </button>

              <button
                onClick={() => setDisplayMode('hp')}
                title="Tampilan Mode HP (Smartphone): Tampilan ramah ponsel, drawer mobile & navigasi sentuh"
                aria-label="Mode Tampilan HP"
                className={`w-7 h-7 inline-flex items-center justify-center rounded-md transition-all cursor-pointer ${
                  displayMode === 'hp'
                    ? 'bg-emerald-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white hover:bg-slate-700/60'
                }`}
              >
                <Smartphone size={14} />
              </button>

              <button
                onClick={() => setDisplayMode('full')}
                title="Tampilan Mode FULL (Layar Penuh): Ruang kerja maksimal tanpa batas tepi untuk tabel nilai"
                aria-label="Mode Tampilan Penuh"
                className={`w-7 h-7 inline-flex items-center justify-center rounded-md transition-all cursor-pointer ${
                  displayMode === 'full'
                    ? 'bg-purple-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white hover:bg-slate-700/60'
                }`}
              >
                <Maximize2 size={14} />
              </button>

              {/* Fullscreen API Toggle Button */}
              <button
                onClick={toggleFullscreen}
                title={isFullscreen ? 'Keluar dari Layar Penuh Browser' : 'Aktifkan Layar Penuh Browser'}
                aria-label={isFullscreen ? 'Keluar Layar Penuh' : 'Layar Penuh Browser'}
                className="w-7 h-7 inline-flex items-center justify-center rounded-md text-slate-400 hover:text-white hover:bg-slate-700 transition-colors cursor-pointer"
              >
                {isFullscreen ? <Minimize2 size={14} /> : <Maximize2 size={14} />}
              </button>
            </div>

            {/* PENGATURAN BILAH TOMBOL AKSI: PENUH / RAMPING / SEMBUNYI */}
            <div className="flex items-center bg-slate-800/90 p-0.5 rounded-lg border border-slate-700/80 shadow-xs" title="Atur tampilan bilah tombol aksi pada semua tabel data">
              <button
                onClick={() => setActionToolbarMode('full')}
                title="Tampilkan tombol aksi penuh & lengkap di semua menu data"
                aria-label="Bilah Aksi Penuh"
                className={`w-7 h-7 inline-flex items-center justify-center rounded-md transition-all cursor-pointer ${
                  actionToolbarMode === 'full'
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white hover:bg-slate-700/60'
                }`}
              >
                <SlidersHorizontal size={13} />
              </button>

              <button
                onClick={() => setActionToolbarMode('compact')}
                title="Rampingkan bilah tombol aksi untuk menghemat ruang kerja"
                aria-label="Bilah Aksi Ramping"
                className={`w-7 h-7 inline-flex items-center justify-center rounded-md transition-all cursor-pointer ${
                  actionToolbarMode === 'compact'
                    ? 'bg-amber-400 text-slate-950 font-black shadow-xs'
                    : 'text-slate-300 hover:text-white hover:bg-slate-700/60'
                }`}
              >
                <Minimize2 size={13} />
              </button>

              <button
                onClick={() => setActionToolbarMode('hidden')}
                title="Sembunyikan bilah tombol aksi untuk ruang kerja data maksimal"
                aria-label="Sembunyikan Bilah Aksi"
                className={`w-7 h-7 inline-flex items-center justify-center rounded-md transition-all cursor-pointer ${
                  actionToolbarMode === 'hidden'
                    ? 'bg-rose-600 text-white shadow-xs'
                    : 'text-slate-300 hover:text-white hover:bg-slate-700/60'
                }`}
              >
                <EyeOff size={13} />
              </button>
            </div>

            {/* Tombol Simpan Semua Data & Status Penyimpanan */}
            <div className="inline-flex items-center gap-1">
              <button
                onClick={saveAllData}
                disabled={globalSaveStatus === 'saving'}
                title="Simpan seluruh data ke database lokal (Ctrl+S)"
                aria-label="Simpan Semua Data"
                className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white transition-all shadow-xs cursor-pointer border border-emerald-400 active:scale-95 disabled:opacity-60 disabled:cursor-not-allowed"
              >
                <Save size={14} />
              </button>
              <SaveStatusIndicator
                status={globalSaveStatus}
                compact
                lastSavedText={lastSavedTime}
              />
            </div>

            {/* Tombol Cloud Sync Firebase Firestore */}
            <button
              onClick={() => syncToCloudFirestore()}
              disabled={cloudSyncStatus === 'syncing'}
              title={
                cloudSyncStatus === 'syncing'
                  ? 'Menyinkronkan ke Cloud Firebase Firestore...'
                  : lastCloudSyncTime
                  ? `Sinkronkan ke Cloud Firebase Firestore (Terakhir: ${lastCloudSyncTime})`
                  : 'Sinkronkan data ke Cloud Firebase Firestore'
              }
              aria-label="Sinkronisasi Cloud"
              className={`w-7 h-7 inline-flex items-center justify-center rounded-lg transition-all shadow-xs cursor-pointer border active:scale-95 disabled:opacity-60 ${
                cloudSyncStatus === 'synced'
                  ? 'bg-cyan-700 hover:bg-cyan-600 text-white border-cyan-400'
                  : 'bg-indigo-600 hover:bg-indigo-500 text-white border-indigo-400'
              }`}
            >
              <CloudUpload size={14} className={cloudSyncStatus === 'syncing' ? 'animate-bounce' : ''} />
            </button>

            {/* Tombol Pintas Google Sheets */}
            <button
              onClick={() => setActiveMenu('google_sheets')}
              title="Buka Integrasi Google Sheets & Google Drive (Impor / Ekspor Data)"
              aria-label="Google Sheets"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-emerald-700 hover:bg-emerald-600 text-white transition-all shadow-xs cursor-pointer border border-emerald-400 active:scale-95"
            >
              <FileSpreadsheet size={14} className="text-emerald-200" />
            </button>

            {/* Tombol Sinkronkan Hubungan Semua Data */}
            <button
              onClick={() => syncAllData()}
              title="Hubungkan & Sinkronkan Semua Data yang Relevan (Sekolah, Siswa, Nilai, Agama, Rapor, Kepala Sekolah, Wali Kelas)"
              aria-label="Sinkronkan Semua Data"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-amber-400 hover:bg-amber-300 text-slate-950 transition-all shadow-xs cursor-pointer border border-amber-300 active:scale-95"
            >
              <Zap size={14} className="text-slate-950" />
            </button>

            {/* Tombol Bersihkan Data Semua di Semua Menu */}
            <button
              type="button"
              onClick={openClearDataModal}
              title="Bersihkan Data Semua di Semua Menu atau Data Menu Aktif"
              aria-label="Bersihkan Data Semua"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-rose-600 hover:bg-rose-500 text-white transition-all shadow-xs cursor-pointer border border-rose-400 active:scale-95"
            >
              <Trash2 size={14} />
            </button>

            {/* Status Jaringan Online / Offline */}
            <button
              type="button"
              onClick={() => setShowInstallModal(true)}
              title={
                isOnline
                  ? isOfflineReady
                    ? 'Online & Cache Offline Siap. Klik untuk Instal / Mode Offline'
                    : 'Online. Klik untuk Instal di Laptop / Unduh Offline'
                  : 'Mode Offline Aktif & Siap Digunakan'
              }
              aria-label={isOnline ? 'Status Online' : 'Status Offline'}
              className={`w-7 h-7 inline-flex items-center justify-center rounded-lg border transition-all cursor-pointer relative ${
                isOnline
                  ? 'bg-emerald-950/90 hover:bg-emerald-900 border-emerald-500/50 text-emerald-300'
                  : 'bg-amber-950/90 hover:bg-amber-900 border-amber-500/60 text-amber-300'
              }`}
            >
              {isOnline ? (
                <Wifi size={14} className="text-emerald-400" />
              ) : (
                <WifiOff size={14} className="text-amber-400" />
              )}
              <span
                className={`absolute top-1 right-1 w-1.5 h-1.5 rounded-full ${
                  isOnline ? 'bg-emerald-400' : 'bg-amber-400 animate-ping'
                }`}
              />
            </button>

            {/* Developer badge button */}
            <button
              onClick={() => setShowDevModal(true)}
              title="Profil Pengembang: Selpius Nabyal, S.Pd.Gr."
              aria-label="Profil Pengembang"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-emerald-950/80 hover:bg-emerald-900 border border-emerald-500/40 text-emerald-200 transition-all shadow-xs cursor-pointer group"
            >
              <Award size={14} className="text-amber-400 group-hover:scale-110 transition-transform" />
            </button>

            {/* Tombol Instal PC/Laptop & Mode Offline (PWA) */}
            <PWAInstallButton onOpenInstallGuide={() => setShowInstallModal(true)} />

            {/* Tombol Bagikan */}
            <button
              onClick={() => setShowShareModal(true)}
              title="Bagikan tautan aplikasi ke rekan guru atau instal di laptop lain"
              aria-label="Bagikan Aplikasi"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white shadow-xs transition-all cursor-pointer ring-1 ring-emerald-400/40"
            >
              <Share2 size={14} />
            </button>

            {/* Tombol Preview Rapor Langsung */}
            <button
              onClick={() => openReportPreview()}
              title="Buka Pratinjau Rapor Lengkap Siswa (Kaver, Sekolah, Siswa, Nilai & TTD)"
              aria-label="Preview Rapor"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-gradient-to-r from-purple-600 via-indigo-600 to-blue-600 hover:from-purple-500 hover:via-indigo-500 hover:to-blue-500 text-white shadow-xs transition-all cursor-pointer ring-1 ring-purple-400/50"
            >
              <Eye size={14} className="text-amber-300" />
            </button>

            {/* Tombol Buku Panduan */}
            <button
              onClick={() => setActiveMenu('buku_panduan')}
              title="Buka Buku Panduan Lengkap Aplikasi"
              aria-label="Buku Panduan"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-indigo-600 hover:bg-indigo-500 text-white shadow-xs transition-all cursor-pointer ring-1 ring-indigo-400/40"
            >
              <BookOpen size={14} className="text-amber-300" />
            </button>

            {/* Tombol Cepat Ukuran Huruf UI */}
            <button
              onClick={() => {
                if (uiFontSize === 'small') setUiFontSize('normal');
                else if (uiFontSize === 'normal') setUiFontSize('large');
                else if (uiFontSize === 'large') setUiFontSize('xlarge');
                else setUiFontSize('small');
              }}
              title={`Ukuran Huruf Aplikasi: ${uiFontScalePercent}%. Klik untuk mengubah`}
              aria-label="Ukuran Huruf Aplikasi"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors cursor-pointer"
            >
              <Type size={14} className="text-emerald-400" />
            </button>

            {/* Tombol Cepat Ganti Tema Gelap / Terang */}
            <button
              onClick={() => setThemeMode(themeMode === 'dark' ? 'light' : 'dark')}
              title={themeMode === 'dark' ? 'Ganti ke Mode Terang (Light Mode)' : 'Ganti ke Mode Gelap (Dark Mode)'}
              aria-label="Ganti Tema"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-slate-800 hover:bg-slate-700 text-amber-300 hover:text-amber-200 border border-slate-700 transition-colors cursor-pointer"
            >
              {themeMode === 'dark' ? <Sun size={14} /> : <Moon size={14} />}
            </button>

            {/* Tombol Pintas Pengaturan Lengkap */}
            <button
              onClick={() => setActiveMenu('pengaturan')}
              title="Buka Menu Pengaturan: Huruf, Kertas, Print, dan Tema"
              aria-label="Pengaturan"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white border border-slate-700 transition-colors cursor-pointer"
            >
              <Settings size={14} />
            </button>

            {/* Tombol Tutup Aplikasi */}
            <button
              onClick={() => setShowCloseModal(true)}
              title="Tutup Aplikasi Rapor"
              aria-label="Tutup Aplikasi"
              className="w-7 h-7 inline-flex items-center justify-center rounded-lg bg-rose-950/70 hover:bg-rose-900 border border-rose-700/50 text-rose-200 transition-colors cursor-pointer"
            >
              <Power size={14} className="text-rose-400" />
            </button>
          </div>
        </div>

        {/* Bottom bar: Filter Bar (Kelas, Semester, Fase) & Tombol Masuk/Keluar */}
        <div className="px-4 py-2 bg-slate-900 flex flex-wrap items-center justify-between gap-3 text-xs">
          {/* Filters */}
          <div className="flex flex-wrap items-center gap-2.5">
            <span className="text-slate-400 font-semibold flex items-center gap-1">
              <Sparkles size={13} className="text-amber-400" />
              Filter Utama:
            </span>

            {/* Filter Kelas */}
            <div className="inline-flex items-center gap-1.5 bg-slate-800/90 border border-slate-700 px-2.5 py-1 rounded-lg">
              <span className="text-slate-400 font-medium">Kelas:</span>
              <select
                value={selectedKelas}
                onChange={(e) => setSelectedKelas(e.target.value as ClassLevel)}
                className="bg-transparent text-amber-300 font-bold focus:outline-hidden cursor-pointer"
              >
                <optgroup label="── SD / MI (Kelas 1-6) ──" className="bg-slate-900 text-amber-400 font-bold">
                  <option value="I" className="bg-slate-800 text-white">Kelas I (Fase A)</option>
                  <option value="II" className="bg-slate-800 text-white">Kelas II (Fase A)</option>
                  <option value="III" className="bg-slate-800 text-white">Kelas III (Fase B)</option>
                  <option value="IV" className="bg-slate-800 text-white">Kelas IV (Fase B)</option>
                  <option value="V" className="bg-slate-800 text-white">Kelas V (Fase C)</option>
                  <option value="VI" className="bg-slate-800 text-white">Kelas VI (Fase C)</option>
                </optgroup>
                <optgroup label="── SMP / MTs (Kelas 7-9) ──" className="bg-slate-900 text-blue-400 font-bold">
                  <option value="VII" className="bg-slate-800 text-white">Kelas VII (SMP - Fase D)</option>
                  <option value="VIII" className="bg-slate-800 text-white">Kelas VIII (SMP - Fase D)</option>
                  <option value="IX" className="bg-slate-800 text-white">Kelas IX (SMP - Fase D)</option>
                </optgroup>
                <optgroup label="── SMA / MA & SMK (Kelas 10-12/13) ──" className="bg-slate-900 text-emerald-400 font-bold">
                  <option value="X" className="bg-slate-800 text-white">Kelas X (SMA/SMK - Fase E)</option>
                  <option value="XI" className="bg-slate-800 text-white">Kelas XI (SMA/SMK - Fase F)</option>
                  <option value="XII" className="bg-slate-800 text-white">Kelas XII (SMA/SMK - Fase F)</option>
                  <option value="XIII" className="bg-slate-800 text-white">Kelas XIII (SMK 4 Tahun - Fase F)</option>
                </optgroup>
              </select>
            </div>

            {/* Filter Semester */}
            <div className="inline-flex items-center gap-1.5 bg-slate-800/90 border border-slate-700 px-2.5 py-1 rounded-lg">
              <Calendar size={13} className="text-slate-400" />
              <span className="text-slate-400 font-medium">Semester:</span>
              <select
                value={selectedSemester}
                onChange={(e) => setSelectedSemester(e.target.value as SemesterType)}
                className="bg-transparent text-emerald-300 font-bold focus:outline-hidden cursor-pointer"
              >
                <option value="Ganjil" className="bg-slate-800 text-white">Semester 1 (Ganjil)</option>
                <option value="Genap" className="bg-slate-800 text-white">Semester 2 (Genap)</option>
              </select>
            </div>

            {/* Fase selector */}
            <div className="inline-flex items-center gap-1.5 bg-slate-800/90 border border-slate-700 px-2.5 py-1 rounded-lg">
              <Layers size={13} className="text-cyan-400" />
              <span className="text-slate-400 font-medium">Fase:</span>
              <select
                value={selectedFase}
                onChange={(e) => setSelectedFase(e.target.value as any)}
                className="bg-transparent text-cyan-300 font-bold focus:outline-hidden cursor-pointer"
                title="Pilih Fase Kurikulum Merdeka"
              >
                <option value="A" className="bg-slate-800 text-white">Fase A (SD Kelas I-II)</option>
                <option value="B" className="bg-slate-800 text-white">Fase B (SD Kelas III-IV)</option>
                <option value="C" className="bg-slate-800 text-white">Fase C (SD Kelas V-VI)</option>
                <option value="D" className="bg-slate-800 text-white">Fase D (SMP Kelas VII-IX)</option>
                <option value="E" className="bg-slate-800 text-white">Fase E (SMA/SMK Kelas X)</option>
                <option value="F" className="bg-slate-800 text-white">Fase F (SMA/SMK Kelas XI-XII/XIII)</option>
              </select>
            </div>

            {/* Filter Nama Siswa */}
            <div className="inline-flex items-center gap-1.5 bg-slate-800/90 border border-slate-700 px-2.5 py-1 rounded-lg">
              <Users size={13} className="text-emerald-400" />
              <span className="text-slate-400 font-medium">Siswa:</span>
              <select
                value={selectedStudentFilter}
                onChange={(e) => setSelectedStudentFilter(e.target.value)}
                className="bg-transparent text-emerald-300 font-bold focus:outline-hidden cursor-pointer max-w-[180px]"
                title="Filter Nama Siswa di seluruh menu"
              >
                <option value="SEMUA" className="bg-slate-800 text-white">
                  Semua Siswa ({db.siswa.filter((s) => s.kelas === selectedKelas).length})
                </option>
                {db.siswa
                  .filter((s) => s.kelas === selectedKelas)
                  .map((s) => (
                    <option key={s.idSiswa} value={s.idSiswa} className="bg-slate-800 text-white">
                      {s.namaSiswa}
                    </option>
                  ))}
              </select>
              {selectedStudentFilter !== 'SEMUA' && (
                <button
                  type="button"
                  onClick={() => setSelectedStudentFilter('SEMUA')}
                  title="Reset Filter Nama Siswa"
                  className="text-slate-400 hover:text-rose-300 cursor-pointer"
                >
                  <X size={12} />
                </button>
              )}
            </div>
          </div>

          {/* User Status & Login/Logout buttons */}
          <div className="flex items-center gap-2 ml-auto">
            {currentUser && (
              <div className="flex items-center gap-2 bg-slate-800/80 px-2.5 py-1 rounded-lg border border-slate-700">
                <div className="w-5 h-5 rounded-full bg-emerald-700 flex items-center justify-center text-[10px] font-bold text-white">
                  {currentUser.namaLengkap.charAt(0)}
                </div>
                <div className="text-left hidden lg:block">
                  <span className="text-[11px] font-semibold text-slate-200 block leading-tight">
                    {currentUser.namaLengkap}
                  </span>
                  <span className="text-[9px] text-emerald-400 block leading-tight">
                    {currentUser.role}
                  </span>
                </div>
                <button
                  onClick={logout}
                  title="Keluar dari akun"
                  className="ml-1 p-1 hover:bg-slate-700 text-slate-400 hover:text-rose-400 rounded transition-colors cursor-pointer"
                >
                  <LogOut size={14} />
                </button>
              </div>
            )}
            <button
              type="button"
              onClick={() => setShowAuthModal(true)}
              title="Buka Form Login Aplikasi Rapor"
              aria-label="Login Aplikasi Rapor"
              className="w-8 h-8 inline-flex items-center justify-center bg-black hover:bg-neutral-900 text-[#00FFFF] border-2 border-t-[#f0f0f0] border-l-[#f0f0f0] border-r-[#7a7a7a] border-b-[#7a7a7a] rounded-sm transition-all shadow-sm cursor-pointer active:translate-y-px"
            >
              <LogIn size={14} className="text-[#00FFFF]" />
            </button>
          </div>
        </div>
      </header>

      {/* Modals */}
      <DeveloperModal isOpen={showDevModal} onClose={() => setShowDevModal(false)} />
      <InstallPcModal isOpen={showInstallModal} onClose={() => setShowInstallModal(false)} />
      <ShareModal isOpen={showShareModal} onClose={() => setShowShareModal(false)} />
      <CloseAppModal isOpen={showCloseModal} onClose={() => setShowCloseModal(false)} />
      <LoginRegisterModal isOpen={showAuthModal} onClose={() => setShowAuthModal(false)} />
    </>
  );
};
