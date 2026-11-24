import React, { useRef, useState } from 'react';
import {
  Save,
  RotateCcw,
  RotateCw,
  RefreshCw,
  FileUp,
  FileDown,
  Trash2,
  X,
  Edit3,
  Upload,
  Eraser,
  CheckCircle,
  FileSpreadsheet,
  Printer,
  Copy,
  Eye,
  EyeOff,
  Minimize2,
  Maximize2,
  ChevronDown,
  ChevronUp,
  SlidersHorizontal
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { SaveStatusIndicator, SaveStatus } from './SaveStatusIndicator';
import { useSaveFeedback } from '../../hooks/useSaveFeedback';
import { parseAndImportTemplate } from '../../utils/templateImporter';

export interface ActionToolbarProps {
  title?: string;
  badge?: string;

  // 1. Simpan
  onSave?: () => void;
  saveLabel?: string;
  saveDisabled?: boolean;
  hideSave?: boolean;
  saveStatus?: SaveStatus;

  // 2. Update
  onUpdate?: () => void;
  updateLabel?: string;
  updateDisabled?: boolean;
  hideUpdate?: boolean;

  // 3. Edit
  onEdit?: () => void;
  editLabel?: string;
  editDisabled?: boolean;
  isEditing?: boolean;
  hideEdit?: boolean;

  // 4. Batal
  onCancel?: () => void;
  cancelLabel?: string;
  hideCancel?: boolean;

  // 5. Hapus
  onDelete?: () => void;
  deleteLabel?: string;
  deleteDisabled?: boolean;
  hideDelete?: boolean;

  // 6. Impor
  onImport?: (file: File) => void;
  importAccept?: string;
  importLabel?: string;
  hideImport?: boolean;

  // 7. Expor
  onExport?: () => void;
  exportLabel?: string;
  hideExport?: boolean;

  // 8. Upload
  onUpload?: (file: File) => void;
  uploadAccept?: string;
  uploadLabel?: string;
  hideUpload?: boolean;

  // 9. Bersihkan File / Input
  onClear?: () => void;
  clearLabel?: string;
  hideClear?: boolean;

  // 10. Undo
  onUndo?: () => void;
  hideUndo?: boolean;

  // 11. Redo
  onRedo?: () => void;
  hideRedo?: boolean;

  // 12. Refres
  onRefresh?: () => void;
  hideRefresh?: boolean;

  // 13. Template (Aksi Tambahan Relevan)
  onTemplate?: () => void;
  templateLabel?: string;
  hideTemplate?: boolean;

  // 14. Cetak (Aksi Tambahan Relevan)
  onPrint?: () => void;
  hidePrint?: boolean;

  // 15. Salin (Aksi Tambahan Relevan)
  onCopy?: () => void;
  hideCopy?: boolean;

  // 16. Preview Rapor
  onPreview?: () => void;
  previewLabel?: string;
  hidePreview?: boolean;

  selectedCount?: number;
  className?: string;
  compact?: boolean;
}

export const StandardActionToolbar: React.FC<ActionToolbarProps> = ({
  title,
  badge,
  onSave,
  saveLabel = 'Simpan',
  saveDisabled = false,
  hideSave = false,
  saveStatus: externalSaveStatus,

  onUpdate,
  updateLabel = 'Update',
  updateDisabled = false,
  hideUpdate = false,

  onEdit,
  editLabel = 'Edit',
  editDisabled = false,
  isEditing = false,
  hideEdit = false,

  onCancel,
  cancelLabel = 'Batal',
  hideCancel = false,

  onDelete,
  deleteLabel = 'Hapus',
  deleteDisabled = false,
  hideDelete = false,

  onImport,
  importAccept = '.xlsx,.xls,.csv',
  importLabel = 'Impor',
  hideImport = false,

  onExport,
  exportLabel = 'Expor',
  hideExport = false,

  onUpload,
  uploadAccept = '*/*',
  uploadLabel = 'Upload',
  hideUpload = false,

  onClear,
  clearLabel = 'Bersihkan File',
  hideClear = false,

  onUndo,
  hideUndo = false,

  onRedo,
  hideRedo = false,

  onRefresh,
  hideRefresh = false,

  onTemplate,
  templateLabel = 'Template',
  hideTemplate = false,

  onPrint,
  hidePrint = false,

  onCopy,
  hideCopy = false,

  onPreview,
  previewLabel = 'Preview Rapor',
  hidePreview = false,

  selectedCount,
  className = '',
  compact = false
}) => {
  const appContext = useApp();
  const {
    undo,
    redo,
    canUndo,
    canRedo,
    refreshData,
    showSaveNotification,
    setActiveMenu,
    saveAllData,
    openReportPreview,
    openClearDataModal,
    actionToolbarMode,
    setActionToolbarMode
  } = appContext;

  const importFileRef = useRef<HTMLInputElement>(null);
  const uploadFileRef = useRef<HTMLInputElement>(null);
  const [internalEditing, setInternalEditing] = useState(false);
  const [localSecondaryExpanded, setLocalSecondaryExpanded] = useState(false);

  // Hook indikator status simpan lokal (Saving..., Saved)
  const { status: internalSaveStatus, triggerSave, lastSavedTime } = useSaveFeedback();
  const activeSaveStatus = externalSaveStatus || internalSaveStatus;

  const activeIsEditing = isEditing || internalEditing;

  // Tentukan mode efektif (full | compact | hidden)
  const effectiveMode = actionToolbarMode === 'hidden'
    ? 'hidden'
    : (actionToolbarMode === 'compact' || compact ? 'compact' : 'full');

  const isCompactMode = effectiveMode === 'compact';

  const handleSaveClick = () => {
    triggerSave(() => {
      if (onSave) {
        onSave();
      } else {
        saveAllData();
      }
    });
  };

  const handleUpdateClick = () => {
    if (onUpdate) {
      onUpdate();
    } else if (onSave) {
      onSave();
    } else {
      showSaveNotification('Data berhasil diperbarui!');
    }
  };

  const handleEditClick = () => {
    if (onEdit) {
      onEdit();
    } else {
      setInternalEditing(!internalEditing);
      showSaveNotification(!internalEditing ? 'Mode Edit diaktifkan!' : 'Mode Edit dinonaktifkan.');
    }
  };

  const handleCancelClick = () => {
    if (onCancel) {
      onCancel();
    } else {
      setInternalEditing(false);
      showSaveNotification('Perubahan dibatalkan / pilihan direset.');
    }
  };

  const handleDeleteClick = () => {
    if (onDelete) {
      onDelete();
    } else {
      showSaveNotification('Pilih data terlebih dahulu untuk dihapus.');
    }
  };

  const handleImportFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (onImport) {
        onImport(file);
      } else {
        try {
          const res = await parseAndImportTemplate(file, undefined, appContext);
          if (res.success) {
            showSaveNotification(res.message);
          } else {
            showSaveNotification(`Gagal Impor: ${res.message}`);
          }
        } catch (err: any) {
          showSaveNotification(`Gagal membaca file: ${err.message || 'Format tidak dikenali'}`);
        }
      }
    }
    if (e.target) e.target.value = '';
  };

  const handleUploadFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (onUpload) {
        onUpload(file);
      } else {
        try {
          const res = await parseAndImportTemplate(file, undefined, appContext);
          if (res.success) {
            showSaveNotification(res.message);
          } else {
            showSaveNotification(`File lampiran "${file.name}" berhasil diunggah!`);
          }
        } catch (err: any) {
          showSaveNotification(`Gagal membaca file: ${err.message || 'Format tidak dikenali'}`);
        }
      }
    }
    if (e.target) e.target.value = '';
  };

  const handleExportClick = () => {
    if (onExport) {
      onExport();
    } else {
      showSaveNotification('Mengekspor data tabel ke format Excel (.xlsx)...');
    }
  };

  const handleClearClick = () => {
    if (onClear) {
      onClear();
    } else {
      if (importFileRef.current) importFileRef.current.value = '';
      if (uploadFileRef.current) uploadFileRef.current.value = '';
      showSaveNotification('Input dan file berhasil dibersihkan.');
    }
  };

  const handleUndo = () => {
    if (onUndo) onUndo();
    else undo();
  };

  const handleRedo = () => {
    if (onRedo) onRedo();
    else redo();
  };

  const handleRefresh = () => {
    if (onRefresh) onRefresh();
    else refreshData();
  };

  const handleTemplateClick = () => {
    if (onTemplate) onTemplate();
    else setActiveMenu('template_data');
  };

  const handlePrintClick = () => {
    if (onPrint) onPrint();
    else window.print();
  };

  const handleCopyClick = () => {
    if (onCopy) {
      onCopy();
    } else {
      showSaveNotification('Data tabel berhasil disalin ke clipboard!');
    }
  };

  const handlePreviewClick = () => {
    if (onPreview) {
      onPreview();
    } else {
      openReportPreview();
    }
  };

  // Hitung jumlah tombol sekunder yang aktif tersedia
  const availableSecondaryCount = [
    !hideImport,
    !hideUpload,
    !hideClear,
    !hideUndo,
    !hideRedo,
    !hideRefresh,
    !hideTemplate,
    !hideCopy,
    !hideCancel,
    !hideDelete
  ].filter(Boolean).length;

  // ─────────────────────────────────────────────────────────────
  // 1. TAMPILAN JIKA MODE SEMBUNYI (HIDDEN)
  // Menampilkan bilah ramping minimalis dengan tombol ikon saja
  // ─────────────────────────────────────────────────────────────
  if (effectiveMode === 'hidden') {
    return (
      <div
        className={`bg-slate-50/95 hover:bg-slate-100/95 border border-dashed border-slate-300 rounded-xl px-3 py-1.5 flex flex-wrap items-center justify-between gap-2 no-print transition-all shadow-xs ${className}`}
      >
        <div className="flex items-center gap-2">
          {title ? (
            <span className="text-xs font-black text-slate-700 uppercase tracking-wide">
              {title}
            </span>
          ) : (
            <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">
              Menu Data
            </span>
          )}
          {badge && (
            <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-blue-100 text-blue-900 border border-blue-200">
              {badge}
            </span>
          )}
          {selectedCount !== undefined && selectedCount > 0 && (
            <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-950 border border-amber-300">
              {selectedCount} terpilih
            </span>
          )}
        </div>

        {/* Tombol Ikon Cepat di Kanan saat Sembunyi */}
        <div className="ml-auto flex items-center justify-end gap-1.5">
          {!hideSave && (
            <button
              type="button"
              onClick={handleSaveClick}
              disabled={saveDisabled || activeSaveStatus === 'saving'}
              className="inline-flex items-center justify-center w-7 h-7 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg cursor-pointer transition-all shadow-xs disabled:opacity-50"
              title={`${saveLabel} (Ctrl+S)`}
              aria-label={saveLabel}
            >
              <Save size={14} />
            </button>
          )}

          <button
            type="button"
            onClick={openClearDataModal}
            className="inline-flex items-center justify-center w-7 h-7 bg-rose-600 hover:bg-rose-700 text-white rounded-lg shadow-xs transition-all cursor-pointer"
            title="Bersihkan Data Semua"
            aria-label="Bersihkan Data Semua"
          >
            <Trash2 size={14} />
          </button>

          <button
            type="button"
            onClick={() => setActionToolbarMode('full')}
            className="inline-flex items-center justify-center w-7 h-7 bg-blue-600 hover:bg-blue-700 text-white rounded-lg shadow-xs transition-all cursor-pointer"
            title="Tampilkan Tombol Aksi Lengkap"
            aria-label="Tampilkan Tombol Aksi Lengkap"
          >
            <Eye size={14} className="text-amber-300" />
          </button>

          <button
            type="button"
            onClick={() => setActionToolbarMode('compact')}
            className="inline-flex items-center justify-center w-7 h-7 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors cursor-pointer"
            title="Buka Mode Ramping"
            aria-label="Buka Mode Ramping"
          >
            <Minimize2 size={14} />
          </button>
        </div>
      </div>
    );
  }

  // ─────────────────────────────────────────────────────────────
  // 2. TAMPILAN PENUH (FULL) ATAU RAMPING (COMPACT)
  // Semua tombol aksi berupa ikon ringkas di sebelah kanan (ml-auto)
  // ─────────────────────────────────────────────────────────────
  const iconBtnBase =
    'inline-flex items-center justify-center w-8 h-8 rounded-lg transition-all cursor-pointer border shadow-2xs shrink-0';

  return (
    <div
      className={`bg-white ${
        isCompactMode ? 'rounded-xl px-2.5 py-2 shadow-2xs gap-1.5' : 'rounded-2xl px-3 py-2.5 shadow-xs gap-2'
      } border border-slate-200 flex flex-col no-print transition-all ${className}`}
    >
      {/* Hidden File Inputs */}
      <input
        type="file"
        ref={importFileRef}
        onChange={handleImportFileChange}
        accept={importAccept}
        className="hidden"
      />
      <input
        type="file"
        ref={uploadFileRef}
        onChange={handleUploadFileChange}
        accept={uploadAccept}
        className="hidden"
      />

      {/* Baris Utama: Judul/Status di Kiri, Tombol Ikon Aksi di Kanan */}
      <div className="flex flex-wrap items-center justify-between gap-2">
        {/* SISI KIRI: Judul Menu, Badge, Status Edit/Pilihan */}
        <div className="flex flex-wrap items-center gap-1.5">
          {title ? (
            <span className="text-xs font-black text-slate-800 uppercase tracking-wide">
              {title}
            </span>
          ) : (
            <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">
              Aksi Data:
            </span>
          )}
          {badge && (
            <span className="px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-blue-100 text-blue-900 border border-blue-200">
              {badge}
            </span>
          )}
          {selectedCount !== undefined && selectedCount > 0 && (
            <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-950 border border-amber-300">
              {selectedCount} terpilih
            </span>
          )}
          {activeIsEditing && (
            <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-900 border border-emerald-300 animate-pulse">
              Edit Aktif
            </span>
          )}
        </div>

        {/* SISI KANAN: SEMUA TOMBOL AKSI BERUPA IKON RINGKAS */}
        <div className="ml-auto flex items-center justify-end flex-wrap gap-1">
          {/* 1. SIMPAN */}
          {!hideSave && (
            <div className="inline-flex items-center gap-1">
              <button
                type="button"
                onClick={handleSaveClick}
                disabled={saveDisabled || activeSaveStatus === 'saving'}
                className={`${iconBtnBase} ${
                  saveDisabled || activeSaveStatus === 'saving'
                    ? 'bg-slate-100 text-slate-400 cursor-not-allowed border-slate-200'
                    : 'bg-emerald-600 hover:bg-emerald-700 text-white border-emerald-700'
                }`}
                title={`${saveLabel} (Ctrl+S)`}
                aria-label={saveLabel}
              >
                <Save size={15} />
              </button>
              <SaveStatusIndicator
                status={activeSaveStatus}
                compact={true}
                lastSavedText={lastSavedTime}
              />
            </div>
          )}

          {/* 2. UPDATE */}
          {!hideUpdate && (
            <button
              type="button"
              onClick={handleUpdateClick}
              disabled={updateDisabled}
              className={`${iconBtnBase} ${
                updateDisabled
                  ? 'bg-slate-100 text-slate-400 cursor-not-allowed border-slate-200'
                  : 'bg-blue-600 hover:bg-blue-700 text-white border-blue-700'
              }`}
              title={updateLabel}
              aria-label={updateLabel}
            >
              <CheckCircle size={15} />
            </button>
          )}

          {/* 3. EDIT */}
          {!hideEdit && (
            <button
              type="button"
              onClick={handleEditClick}
              disabled={editDisabled}
              className={`${iconBtnBase} ${
                activeIsEditing
                  ? 'bg-amber-400 text-slate-950 border-amber-500 ring-2 ring-amber-300'
                  : 'bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-300'
              }`}
              title={editLabel}
              aria-label={editLabel}
            >
              <Edit3 size={15} />
            </button>
          )}

          {/* 4. BATAL */}
          {(!isCompactMode || activeIsEditing) && !hideCancel && (
            <button
              type="button"
              onClick={handleCancelClick}
              className={`${iconBtnBase} bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300`}
              title={cancelLabel}
              aria-label={cancelLabel}
            >
              <X size={15} />
            </button>
          )}

          {/* 5. HAPUS */}
          {(!isCompactMode || (selectedCount !== undefined && selectedCount > 0)) && !hideDelete && (
            <button
              type="button"
              onClick={handleDeleteClick}
              disabled={deleteDisabled}
              className={`${iconBtnBase} ${
                deleteDisabled
                  ? 'bg-slate-100 text-slate-400 cursor-not-allowed border-slate-200'
                  : 'bg-rose-50 hover:bg-rose-600 hover:text-white text-rose-700 border-rose-300'
              }`}
              title={deleteLabel}
              aria-label={deleteLabel}
            >
              <Trash2 size={15} />
            </button>
          )}

          {/* 7. EXPOR */}
          {!hideExport && (
            <button
              type="button"
              onClick={handleExportClick}
              className={`${iconBtnBase} bg-teal-50 hover:bg-teal-100 text-teal-800 border-teal-200`}
              title={exportLabel}
              aria-label={exportLabel}
            >
              <FileDown size={15} />
            </button>
          )}

          {/* 14. CETAK */}
          {!hidePrint && (
            <button
              type="button"
              onClick={handlePrintClick}
              className={`${iconBtnBase} bg-slate-800 hover:bg-slate-900 text-white border-slate-900`}
              title="Cetak"
              aria-label="Cetak"
            >
              <Printer size={15} />
            </button>
          )}

          {/* 16. PREVIEW RAPOR */}
          {!hidePreview && (
            <button
              type="button"
              onClick={handlePreviewClick}
              className={`${iconBtnBase} bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-500 hover:to-indigo-500 text-white border-purple-500`}
              title={previewLabel}
              aria-label={previewLabel}
            >
              <Eye size={15} className="text-amber-300" />
            </button>
          )}

          {/* BERSIHKAN DATA SEMUA */}
          <button
            type="button"
            onClick={openClearDataModal}
            className={`${iconBtnBase} bg-rose-600 hover:bg-rose-700 text-white border-rose-700`}
            title="Bersihkan Data Semua"
            aria-label="Bersihkan Data Semua"
          >
            <Trash2 size={15} />
          </button>

          {/* AKSI SEKUNDER LANGSUNG DI MODE PENUH */}
          {!isCompactMode && (
            <>
              <div className="h-5 w-px bg-slate-300 mx-0.5 hidden sm:block" />

              {!hideImport && (
                <button
                  type="button"
                  onClick={() => importFileRef.current?.click()}
                  className={`${iconBtnBase} bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border-indigo-200`}
                  title={importLabel}
                  aria-label={importLabel}
                >
                  <FileUp size={15} />
                </button>
              )}

              {!hideUpload && (
                <button
                  type="button"
                  onClick={() => uploadFileRef.current?.click()}
                  className={`${iconBtnBase} bg-purple-50 hover:bg-purple-100 text-purple-700 border-purple-200`}
                  title={uploadLabel}
                  aria-label={uploadLabel}
                >
                  <Upload size={15} />
                </button>
              )}

              {!hideClear && (
                <button
                  type="button"
                  onClick={handleClearClick}
                  className={`${iconBtnBase} bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300`}
                  title={clearLabel}
                  aria-label={clearLabel}
                >
                  <Eraser size={15} />
                </button>
              )}

              <div className="h-5 w-px bg-slate-300 mx-0.5 hidden sm:block" />

              {!hideUndo && (
                <button
                  type="button"
                  onClick={handleUndo}
                  disabled={!canUndo}
                  className={`${iconBtnBase} ${
                    canUndo
                      ? 'bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-300'
                      : 'bg-slate-50 text-slate-400 border-slate-200 cursor-not-allowed'
                  }`}
                  title="Undo (Batalkan perubahan terakhir)"
                  aria-label="Undo"
                >
                  <RotateCcw size={15} />
                </button>
              )}

              {!hideRedo && (
                <button
                  type="button"
                  onClick={handleRedo}
                  disabled={!canRedo}
                  className={`${iconBtnBase} ${
                    canRedo
                      ? 'bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-300'
                      : 'bg-slate-50 text-slate-400 border-slate-200 cursor-not-allowed'
                  }`}
                  title="Redo (Ulangi perubahan)"
                  aria-label="Redo"
                >
                  <RotateCw size={15} />
                </button>
              )}

              {!hideRefresh && (
                <button
                  type="button"
                  onClick={handleRefresh}
                  className={`${iconBtnBase} bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300`}
                  title="Refres (Muat ulang data)"
                  aria-label="Refres"
                >
                  <RefreshCw size={15} />
                </button>
              )}

              {!hideTemplate && (
                <button
                  type="button"
                  onClick={handleTemplateClick}
                  className={`${iconBtnBase} bg-blue-50 hover:bg-blue-100 text-blue-800 border-blue-200`}
                  title={templateLabel}
                  aria-label={templateLabel}
                >
                  <FileSpreadsheet size={15} />
                </button>
              )}

              {!hideCopy && (
                <button
                  type="button"
                  onClick={handleCopyClick}
                  className={`${iconBtnBase} bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-200`}
                  title="Salin data tabel ke clipboard"
                  aria-label="Salin"
                >
                  <Copy size={15} />
                </button>
              )}
            </>
          )}

          {/* MODE RAMPING: TOMBOL BUKA/TUTUP AKSI PENDUKUNG (IKON SAJA) */}
          {isCompactMode && availableSecondaryCount > 0 && (
            <button
              type="button"
              onClick={() => setLocalSecondaryExpanded(!localSecondaryExpanded)}
              className={`${iconBtnBase} ${
                localSecondaryExpanded
                  ? 'bg-indigo-100 text-indigo-900 border-indigo-300'
                  : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border-slate-300'
              }`}
              title={`Aksi Lainnya (${availableSecondaryCount} tombol)`}
              aria-label="Aksi Lainnya"
            >
              {localSecondaryExpanded ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
            </button>
          )}

          <div className="h-5 w-px bg-slate-300 mx-0.5 hidden sm:block" />

          {/* KONTROL BILAH: PERLUAS / RAMPINGKAN & SEMBUNYIKAN (IKON SAJA) */}
          <div className="inline-flex items-center rounded-lg bg-slate-100 p-0.5 border border-slate-300 shadow-2xs gap-0.5">
            {isCompactMode ? (
              <button
                type="button"
                onClick={() => setActionToolbarMode('full')}
                className="inline-flex items-center justify-center w-7 h-7 rounded-md text-blue-700 hover:bg-blue-50 transition-colors cursor-pointer"
                title="Perluas semua tombol aksi"
                aria-label="Perluas"
              >
                <Maximize2 size={14} />
              </button>
            ) : (
              <button
                type="button"
                onClick={() => setActionToolbarMode('compact')}
                className="inline-flex items-center justify-center w-7 h-7 rounded-md text-slate-700 hover:bg-slate-200 transition-colors cursor-pointer"
                title="Rampingkan bilah aksi"
                aria-label="Rampingkan"
              >
                <Minimize2 size={14} />
              </button>
            )}

            <button
              type="button"
              onClick={() => setActionToolbarMode('hidden')}
              className="inline-flex items-center justify-center w-7 h-7 rounded-md text-rose-700 hover:bg-rose-50 transition-colors cursor-pointer"
              title="Sembunyikan bilah aksi"
              aria-label="Hilangkan"
            >
              <EyeOff size={14} />
            </button>
          </div>
        </div>
      </div>

      {/* Tray Ekspansi Sekunder saat Mode Ramping (Ikon Saja) */}
      {isCompactMode && localSecondaryExpanded && (
        <div className="pt-2 mt-1 border-t border-slate-100 flex flex-wrap items-center justify-end gap-1.5 animate-fadeIn bg-slate-50/60 p-1.5 rounded-xl">
          <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider mr-auto px-1">
            Aksi Lainnya:
          </span>

          {!hideImport && (
            <button
              type="button"
              onClick={() => importFileRef.current?.click()}
              className={`${iconBtnBase} bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border-indigo-200`}
              title={importLabel}
              aria-label={importLabel}
            >
              <FileUp size={15} />
            </button>
          )}

          {!hideUpload && (
            <button
              type="button"
              onClick={() => uploadFileRef.current?.click()}
              className={`${iconBtnBase} bg-purple-50 hover:bg-purple-100 text-purple-700 border-purple-200`}
              title={uploadLabel}
              aria-label={uploadLabel}
            >
              <Upload size={15} />
            </button>
          )}

          {!hideClear && (
            <button
              type="button"
              onClick={handleClearClick}
              className={`${iconBtnBase} bg-white hover:bg-slate-100 text-slate-700 border-slate-300`}
              title={clearLabel}
              aria-label={clearLabel}
            >
              <Eraser size={15} />
            </button>
          )}

          {!hideUndo && (
            <button
              type="button"
              onClick={handleUndo}
              disabled={!canUndo}
              className={`${iconBtnBase} ${
                canUndo
                  ? 'bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-300'
                  : 'bg-white text-slate-400 border-slate-200 cursor-not-allowed'
              }`}
              title="Undo"
              aria-label="Undo"
            >
              <RotateCcw size={14} />
            </button>
          )}

          {!hideRedo && (
            <button
              type="button"
              onClick={handleRedo}
              disabled={!canRedo}
              className={`${iconBtnBase} ${
                canRedo
                  ? 'bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-300'
                  : 'bg-white text-slate-400 border-slate-200 cursor-not-allowed'
              }`}
              title="Redo"
              aria-label="Redo"
            >
              <RotateCw size={14} />
            </button>
          )}

          {!hideRefresh && (
            <button
              type="button"
              onClick={handleRefresh}
              className={`${iconBtnBase} bg-white hover:bg-slate-100 text-slate-700 border-slate-300`}
              title="Refres"
              aria-label="Refres"
            >
              <RefreshCw size={14} />
            </button>
          )}

          {!hideTemplate && (
            <button
              type="button"
              onClick={handleTemplateClick}
              className={`${iconBtnBase} bg-blue-50 hover:bg-blue-100 text-blue-800 border-blue-200`}
              title={templateLabel}
              aria-label={templateLabel}
            >
              <FileSpreadsheet size={14} />
            </button>
          )}

          {!hideCopy && (
            <button
              type="button"
              onClick={handleCopyClick}
              className={`${iconBtnBase} bg-white hover:bg-slate-100 text-slate-700 border-slate-200`}
              title="Salin data tabel"
              aria-label="Salin"
            >
              <Copy size={14} />
            </button>
          )}

          {!hideCancel && !activeIsEditing && (
            <button
              type="button"
              onClick={handleCancelClick}
              className={`${iconBtnBase} bg-white hover:bg-slate-100 text-slate-700 border-slate-300`}
              title={cancelLabel}
              aria-label={cancelLabel}
            >
              <X size={14} />
            </button>
          )}

          {!hideDelete && !(selectedCount !== undefined && selectedCount > 0) && (
            <button
              type="button"
              onClick={handleDeleteClick}
              disabled={deleteDisabled}
              className={`${iconBtnBase} ${
                deleteDisabled
                  ? 'bg-white text-slate-400 border-slate-200 cursor-not-allowed'
                  : 'bg-rose-50 hover:bg-rose-600 hover:text-white text-rose-700 border-rose-300'
              }`}
              title={deleteLabel}
              aria-label={deleteLabel}
            >
              <Trash2 size={14} />
            </button>
          )}
        </div>
      )}
    </div>
  );
};
