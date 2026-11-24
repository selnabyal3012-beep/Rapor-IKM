import React, { useState, useMemo, useRef, useEffect } from 'react';
import {
  Target,
  Plus,
  Edit2,
  Trash2,
  CheckCircle2,
  Save,
  X,
  Sparkles,
  Search,
  BookOpen,
  FileSpreadsheet,
  Copy,
  Table as TableIcon,
  Download,
  Upload,
  Info,
  Zap,
  Home,
  HelpCircle,
  LayoutGrid
} from 'lucide-react';
import { useApp } from '../../context/AppContext';
import { LearningObjective, ClassLevel, FaseType } from '../../types';
import { SaveStatusIndicator } from '../common/SaveStatusIndicator';
import { StudentFilterSelect } from '../common/StudentFilterSelect';
import { useSaveFeedback } from '../../hooks/useSaveFeedback';
import {
  resolveTpSlots,
  getMatchedTpsForSubjectAndStudent,
  LM_COLUMN_KEYS,
  TP_COLUMN_KEYS
} from '../../utils/tpDescriptionHelper';
import * as XLSX from 'xlsx';
import { parseAndImportTemplate } from '../../utils/templateImporter';

export const ENUM_6_AGAMA = [
  { key: 'Islam', label: 'ISLAM', fullLabel: 'Islam (PAI)', btnClass: 'bg-purple-700 hover:bg-purple-800 text-white border-purple-300' },
  { key: 'Kristen', label: 'KRISTEN', fullLabel: 'Kristen Protestan (PAK)', btnClass: 'bg-blue-600 hover:bg-blue-700 text-white border-blue-300' },
  { key: 'Katolik', label: 'KATOLIK', fullLabel: 'Katolik (PAKAT)', btnClass: 'bg-slate-200 hover:bg-slate-300 text-slate-900 border-slate-400' },
  { key: 'Hindu', label: 'HINDU', fullLabel: 'Hindu (PAH)', btnClass: 'bg-fuchsia-500 hover:bg-fuchsia-600 text-white border-fuchsia-300' },
  { key: 'Buddha', label: 'BUDHA', fullLabel: 'Buddha (PAB)', btnClass: 'bg-emerald-500 hover:bg-emerald-600 text-white border-emerald-300' },
  { key: 'Khonghucu', label: 'KHONGHUCU', fullLabel: 'Khonghucu (PAKh)', btnClass: 'bg-yellow-400 hover:bg-yellow-500 text-slate-950 border-yellow-200' },
  { key: 'Semua', label: 'SEMUA AGAMA', fullLabel: 'Semua Agama (Umum)', btnClass: 'bg-cyan-100 hover:bg-cyan-200 text-slate-950 border-cyan-400' }
] as const;

export const STANDARD_AGAMA_TPS: Record<string, string[]> = {
  Islam: [
    'Siswa mampu membaca dan menghafal surat-surat pendek Al-Qur\'an serta memahami pesan pokok Asmaul Husna dengan tartil.',
    'Siswa mampu mempraktikkan tata cara wudhu dan salat wajib serta menerapkan akhlak mulia kepada orang tua dan sesama manusia.'
  ],
  Kristen: [
    'Siswa mampu memahami kasih Allah yang menciptakan, memelihara, dan menyelamatkan manusia serta segenap alam semesta.',
    'Siswa mampu menerapkan sikap mengasihi sesama, hidup rukun, bersyukur, dan meneladani teladan Tuhan Yesus dalam keseharian.'
  ],
  Katolik: [
    'Siswa mampu memahami peranan Kitab Suci dan Tradisi Gereja dalam meneladani Yesus Kristus sebagai Juru Selamat.',
    'Siswa mampu mewujudkan nilai-nilai Kerajaan Allah melalui perbuatan kasih, doa, dan keterlibatan aktif dalam hidup menggereja.'
  ],
  Hindu: [
    'Siswa mampu memahami konsep Tri Kaya Parisudha dan Catur Asrama dalam menciptakan keharmonisan hidup (Tri Hita Karana).',
    'Siswa mampu mempraktikkan doa Dainika Upasana dan menunjukkan rasa bakti kepada Sang Hyang Widhi Wasa serta orang tua.'
  ],
  Buddha: [
    'Siswa mampu meneladani riwayat hidup Siddharta Gautama serta memahami Empat Kesunyataan Mulia dan Delapan Jalan Utama.',
    'Siswa mampu mempraktikkan sila dan meditasi pernapasan dasar (anapanasati) guna memupuk cinta kasih serta ketenangan batin.'
  ],
  Khonghucu: [
    'Siswa mampu memahami ajaran Xiao (Bakti) kepada orang tua dan mengamalkan Lima Kebajikan Utama (Wu Chang) dalam kehidupan.',
    'Siswa mampu menunjukkan sikap Zhong (Setia) dan Shu (Tepaselira/Tenggang Rasa) serta menghormati sesama makhluk ciptaan Tian.'
  ]
};

const STANDARD_SAMPLE_TPS: Record<string, string[]> = {
  'Pendidikan Agama dan Budi Pekerti': [
    'Siswa mampu memahami ajaran kasih Tuhan, mempraktikkan ibadah sesuai agamanya, dan menunjukkan toleransi antarumat beriman.',
    'Siswa mampu menunjukkan sikap berbudi pekerti luhur, saling menghormati keragaman keyakinan, dan menjaga kerukunan di lingkungan sekolah.'
  ],
  'Pendidikan Pancasila': [
    'Siswa mampu memahami dan menerapkan nilai-nilai Pancasila dalam kegiatan gotong royong di lingkungan sekolah dan keluarga.',
    'Siswa mampu mengenal hak dan kewajiban sebagai peserta didik serta mematuhi aturan musyawarah mufakat di kelas.'
  ],
  'Bahasa Indonesia': [
    'Siswa mampu memahami pesan dan gagasan pokok dari teks narasi maupun eksposisi secara kritis dan mandiri.',
    'Siswa mampu menyampaikan ide, perasaan, dan tanggapan secara lisan maupun tertulis dengan santun dan sistematis.'
  ],
  Matematika: [
    'Siswa mampu membaca, menulis, menentukan nilai tempat, dan membandingkan bilangan cacah sampai 10.000.',
    'Siswa mampu menyelesaikan operasi hitung pecahan sederhana serta menerapkannya dalam pemecahan masalah sehari-hari.'
  ],
  'Ilmu Pengetahuan Alam dan Sosial (IPAS)': [
    'Siswa mampu mengidentifikasi wujud zat, perubahan wujud benda, dan manfaat pemanfaatannya dalam kehidupan sehari-hari.',
    'Siswa mampu menganalisis bentang alam, keanekaragaman hayati, dan kearifan lokal pelestarian lingkungan daerah.'
  ],
  'Pendidikan Jasmani, Olahraga, dan Kesehatan': [
    'Siswa mampu mempraktikkan variasi pola gerak dasar lokomotor, non-lokomotor, dan manipulatif dengan benar.',
    'Siswa mampu menerapkan prinsip pola hidup bersih, sehat, dan menjaga kebugaran jasmani secara konsisten.'
  ],
  'Seni dan Budaya (Seni Rupa & Musik)': [
    'Siswa mampu mengeksplorasi garis, bentuk, tekstur, dan motif ragam hias seni rupa khas kearifan lokal.',
    'Siswa mampu mengapresiasi dan memperagakan irama lagu, gerak berirama, serta ekspresi karya seni sederhana.'
  ],
  'Bahasa Inggris': [
    'Siswa mampu merespons instruksi lisan sederhana dan menggunakan kosakata dasar tentang aktivitas di sekolah dan rumah.',
    'Siswa mampu memperkenalkan diri dan mengungkapkan kesukaan dengan percaya diri menggunakan bahasa Inggris dasar.'
  ],
  'Muatan Lokal (Bahasa Daerah & Noken Papua)': [
    'Siswa mampu mengenal kosakata dasar bahasa suku setempat serta memahami cerita rakyat dan budaya daerah.',
    'Siswa mampu mengapresiasi kearifan budaya pembuatan noken, seni tradisional, dan nilai luhur leluhur setempat.'
  ]
};

const STANDARD_SAMPLE_LMS: Record<string, string[]> = {
  'Pendidikan Agama dan Budi Pekerti': ['Keimanan & Ketakwaan', 'Ibadah & Doa', 'Akhlak Mulia', 'Toleransi & Kerukunan'],
  'Pendidikan Pancasila': ['Nilai Pancasila', 'Aturan & Norma', 'Hak & Kewajiban', 'Bhinneka Tunggal Ika'],
  'Bahasa Indonesia': ['Menyimak & Membaca', 'Kosakata & Kalimat', 'Menulis Narasi', 'Berbicara & Diskusi'],
  Matematika: ['Bilangan Cacah', 'Operasi Hitung', 'Pecahan & Desimal', 'Bangun Datar & Pengukuran'],
  'Ilmu Pengetahuan Alam dan Sosial (IPAS)': ['Wujud Zat & Energi', 'Makhluk Hidup', 'Bentang Alam', 'Kearifan Lokal'],
  'Pendidikan Jasmani, Olahraga, dan Kesehatan': ['Gerak Lokomotor', 'Gerak Manipulatif', 'Kebugaran Jasmani', 'Pola Hidup Sehat'],
  'Seni dan Budaya (Seni Rupa & Musik)': ['Unsur Rupa', 'Karya Kriya Lokal', 'Irama & Lagu', 'Apresiasi Seni'],
  'Bahasa Inggris': ['Greetings & Intro', 'Classroom Objects', 'Daily Activities', 'Simple Sentences'],
  'Muatan Lokal (Bahasa Daerah & Noken Papua)': ['Kosakata Daerah', 'Cerita Rakyat', 'Budaya Noken', 'Adat Istiadat']
};

const SUBJECT_RAIL_COLORS = [
  'bg-amber-200 hover:bg-amber-300 text-slate-950',
  'bg-emerald-200 hover:bg-emerald-300 text-slate-950',
  'bg-sky-200 hover:bg-sky-300 text-slate-950',
  'bg-pink-200 hover:bg-pink-300 text-slate-950',
  'bg-yellow-200 hover:bg-yellow-300 text-slate-950',
  'bg-lime-200 hover:bg-lime-300 text-slate-950',
  'bg-violet-200 hover:bg-violet-300 text-slate-950',
  'bg-orange-200 hover:bg-orange-300 text-slate-950',
  'bg-teal-200 hover:bg-teal-300 text-slate-950'
];

export const LearningObjectiveSection: React.FC = () => {
  const appContext = useApp();
  const {
    db,
    saveTP,
    bulkSaveTujuanPembelajaran,
    deleteTP,
    saveLingkupMateriTitles,
    selectedKelas,
    setSelectedKelas,
    selectedFase,
    selectedTpMapel,
    setSelectedTpMapel,
    filteredStudent,
    syncAllScoresWithTPDescriptions,
    showSaveNotification,
    setActiveMenu
  } = appContext;

  const [viewMode, setViewMode] = useState<'spreadsheet' | 'list_table'>('spreadsheet');
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [selectedAgamaFilter, setSelectedAgamaFilter] = useState<string>('Semua');
  const [showFormulaHelp, setShowFormulaHelp] = useState<boolean>(false);
  const [modalOpen, setModalOpen] = useState<boolean>(false);
  const [editingTP, setEditingTP] = useState<LearningObjective | null>(null);
  const [notification, setNotification] = useState<string | null>(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState<string | null>(null);

  // Draft state for instant 60fps typing on spreadsheet cells
  const [tpDrafts, setTpDrafts] = useState<Record<string, string>>({});
  const [lmDrafts, setLmDrafts] = useState<Record<string, string>>({});
  const debounceTimersRef = useRef<Record<string, number>>({});
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Sync selectedAgamaFilter when a student filter is chosen
  useEffect(() => {
    if (filteredStudent?.agama) {
      setSelectedAgamaFilter(filteredStudent.agama);
    }
  }, [filteredStudent]);

  // Clear local drafts when class or religion filter changes so persisted values show freshly
  useEffect(() => {
    setTpDrafts({});
    setLmDrafts({});
  }, [selectedKelas, selectedAgamaFilter]);

  const showNotice = (msg: string) => {
    setNotification(msg);
    showSaveNotification(msg);
    setTimeout(() => setNotification(null), 3500);
  };

  const defaultNewTP: LearningObjective = {
    id: `TP${String(db.tujuanPembelajaran.length + 1).padStart(3, '0')}`,
    kodeTp: `TP ${db.tujuanPembelajaran.length + 1}`,
    mataPelajaran:
      selectedTpMapel !== 'SEMUA'
        ? selectedTpMapel
        : db.mataPelajaran[0]?.mataPelajaran || 'Bahasa Indonesia',
    fase: selectedFase,
    kelas: selectedKelas,
    agama: selectedAgamaFilter || 'Semua',
    tujuanPembelajaran: ''
  };

  const [formData, setFormData] = useState<LearningObjective>(defaultNewTP);

  const getNextKodeTP = () => {
    const existingCodes = new Set(db.tujuanPembelajaran.map((t) => t.kodeTp.toUpperCase()));
    let nextNum = db.tujuanPembelajaran.length + 1;
    let candidate = `TP ${nextNum}`;
    while (existingCodes.has(candidate.toUpperCase())) {
      nextNum++;
      candidate = `TP ${nextNum}`;
    }
    return candidate;
  };

  const handleOpenAdd = (targetMapel?: string, targetAgama?: string) => {
    setEditingTP(null);
    const chosenMapel =
      targetMapel ||
      (selectedTpMapel !== 'SEMUA'
        ? selectedTpMapel
        : db.mataPelajaran[0]?.mataPelajaran || 'Bahasa Indonesia');
    const nextCode = getNextKodeTP();
    setFormData({
      id: `TP_${Date.now()}`,
      kodeTp: nextCode,
      mataPelajaran: chosenMapel,
      fase: selectedFase,
      kelas: selectedKelas,
      agama: targetAgama || selectedAgamaFilter || 'Semua',
      tujuanPembelajaran: ''
    });
    setModalOpen(true);
  };

  const handleOpenEdit = (tp: LearningObjective) => {
    setEditingTP(tp);
    setFormData({ ...tp, agama: tp.agama || 'Semua' });
    setModalOpen(true);
  };

  const handleDuplicateTP = (tp: LearningObjective) => {
    const nextCode = getNextKodeTP();
    const duplicated: LearningObjective = {
      ...tp,
      id: `TP_${Date.now()}`,
      kodeTp: nextCode,
      tujuanPembelajaran: `${tp.tujuanPembelajaran} (Salinan)`
    };
    saveTP(duplicated);
    showNotice(`Berhasil menyalin TP dengan kode: ${nextCode}`);
  };

  const { status: modalSaveStatus, triggerSave: triggerModalSave } = useSaveFeedback();

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.tujuanPembelajaran.trim()) {
      showNotice('Mohon isi rumusan narasi Tujuan Pembelajaran (TP).');
      return;
    }
    triggerModalSave(async () => {
      saveTP(formData);
      showNotice(
        editingTP
          ? `Tujuan Pembelajaran ${formData.kodeTp} diperbarui & dihubungkan ke deskripsi rapor.`
          : `Tujuan Pembelajaran ${formData.kodeTp} disimpan & dihubungkan ke deskripsi rapor.`
      );
      setTimeout(() => {
        setModalOpen(false);
      }, 500);
    });
  };

  const handleDelete = (id: string, kode: string) => {
    deleteTP(id);
    setConfirmDeleteId(null);
    showNotice(`Tujuan Pembelajaran ${kode} dihapus dan deskripsi rapor diperbarui.`);
  };

  const makeSlug = (val: string) =>
    val
      .replace(/[^a-zA-Z0-9]/g, '')
      .slice(0, 14) || 'MAPEL';

  // Inline Spreadsheet cell update: Lingkup Materi (LM 1..8)
  const commitLmCell = (mapelName: string, lmKey: string, val: string) => {
    const scopeKey = `${selectedKelas}::${mapelName}`;
    const existing = { ...(db.lingkupMateriTitles?.[scopeKey] || {}) };
    if (val.trim()) {
      existing[lmKey] = val;
    } else {
      delete existing[lmKey];
    }
    saveLingkupMateriTitles(scopeKey, existing, true);
  };

  const handleLmChange = (mapelName: string, lmKey: string, val: string) => {
    const draftKey = `${selectedKelas}::${mapelName}::${lmKey}`;
    setLmDrafts((prev) => ({ ...prev, [draftKey]: val }));
    const timerKey = `lm::${draftKey}`;
    if (debounceTimersRef.current[timerKey]) {
      window.clearTimeout(debounceTimersRef.current[timerKey]);
    }
    debounceTimersRef.current[timerKey] = window.setTimeout(() => {
      commitLmCell(mapelName, lmKey, val);
    }, 300);
  };

  const handleLmBlur = (mapelName: string, lmKey: string) => {
    const draftKey = `${selectedKelas}::${mapelName}::${lmKey}`;
    const timerKey = `lm::${draftKey}`;
    if (debounceTimersRef.current[timerKey]) {
      window.clearTimeout(debounceTimersRef.current[timerKey]);
    }
    if (draftKey in lmDrafts) {
      commitLmCell(mapelName, lmKey, lmDrafts[draftKey]);
    }
  };

  // Inline Spreadsheet cell update: Tujuan Pembelajaran (TP 1..12)
  const commitTpCell = (
    mapelName: string,
    slotIdx: number,
    val: string,
    existingTp?: LearningObjective
  ) => {
    const isAgamaSubject = mapelName.toLowerCase().includes('agama');
    const targetAgama = isAgamaSubject ? selectedAgamaFilter || 'Semua' : 'Semua';
    const kodeTp = `TP ${slotIdx + 1}`;
    const targetId =
      existingTp?.id ||
      `TP_${selectedKelas}_${makeSlug(mapelName)}_${makeSlug(targetAgama)}_TP${slotIdx + 1}`;

    saveTP(
      {
        id: targetId,
        kodeTp: existingTp?.kodeTp || kodeTp,
        mataPelajaran: mapelName,
        fase: selectedFase,
        kelas: selectedKelas,
        agama: existingTp?.agama || targetAgama,
        tujuanPembelajaran: val
      },
      true
    );
  };

  const handleTpChange = (
    mapelName: string,
    slotIdx: number,
    val: string,
    existingTp?: LearningObjective
  ) => {
    const isAgamaSubject = mapelName.toLowerCase().includes('agama');
    const agamaKey = isAgamaSubject ? selectedAgamaFilter || 'Semua' : 'Semua';
    const draftKey = `${selectedKelas}::${mapelName}::${agamaKey}::tp${slotIdx + 1}`;
    setTpDrafts((prev) => ({ ...prev, [draftKey]: val }));
    const timerKey = `tp::${draftKey}`;
    if (debounceTimersRef.current[timerKey]) {
      window.clearTimeout(debounceTimersRef.current[timerKey]);
    }
    debounceTimersRef.current[timerKey] = window.setTimeout(() => {
      commitTpCell(mapelName, slotIdx, val, existingTp);
    }, 320);
  };

  const handleTpBlur = (
    mapelName: string,
    slotIdx: number,
    existingTp?: LearningObjective
  ) => {
    const isAgamaSubject = mapelName.toLowerCase().includes('agama');
    const agamaKey = isAgamaSubject ? selectedAgamaFilter || 'Semua' : 'Semua';
    const draftKey = `${selectedKelas}::${mapelName}::${agamaKey}::tp${slotIdx + 1}`;
    const timerKey = `tp::${draftKey}`;
    if (debounceTimersRef.current[timerKey]) {
      window.clearTimeout(debounceTimersRef.current[timerKey]);
    }
    if (draftKey in tpDrafts) {
      commitTpCell(mapelName, slotIdx, tpDrafts[draftKey], existingTp);
    }
  };

  // Seed sample TPs + LMs for a specific subject into TP 1..2 & LM 1..4
  const handleSeedSamplesForMapel = (mapelName: string) => {
    const isAgamaSubject = mapelName.toLowerCase().includes('agama');
    const targetAgama = isAgamaSubject ? selectedAgamaFilter || 'Semua' : 'Semua';
    const samples =
      (isAgamaSubject && targetAgama !== 'Semua' ? STANDARD_AGAMA_TPS[targetAgama] : null) ||
      STANDARD_SAMPLE_TPS[mapelName] || [
        `Siswa mampu memahami materi dan konsep dasar pada pembelajaran ${mapelName} dengan baik.`,
        `Siswa mampu mempraktikkan keterampilan dan menerapkan hasil pembelajaran ${mapelName} dalam kehidupan nyata.`
      ];

    const toSave: LearningObjective[] = samples.map((text, idx) => ({
      id: `TP_${selectedKelas}_${makeSlug(mapelName)}_${makeSlug(targetAgama)}_TP${idx + 1}`,
      kodeTp: `TP ${idx + 1}`,
      mataPelajaran: mapelName,
      fase: selectedFase,
      kelas: selectedKelas,
      agama: targetAgama,
      tujuanPembelajaran: text
    }));
    bulkSaveTujuanPembelajaran(toSave);

    const lmSamples = STANDARD_SAMPLE_LMS[mapelName];
    if (lmSamples) {
      const scopeKey = `${selectedKelas}::${mapelName}`;
      const existingLm = { ...(db.lingkupMateriTitles?.[scopeKey] || {}) };
      lmSamples.forEach((name, idx) => {
        const k = `lm${idx + 1}`;
        if (!existingLm[k]) existingLm[k] = name;
      });
      saveLingkupMateriTitles(scopeKey, existingLm, true);
    }

    showNotice(`Berhasil mengisi contoh TP & Lingkup Materi untuk ${mapelName}.`);
  };

  // Seed TP untuk 6 Agama Resmi di Indonesia
  const handleSeed6Agama = () => {
    const mapelAgama =
      db.mataPelajaran.find((m) => m.mataPelajaran.toLowerCase().includes('agama'))?.mataPelajaran ||
      'Pendidikan Agama dan Budi Pekerti';
    const toSave: LearningObjective[] = [];
    Object.entries(STANDARD_AGAMA_TPS).forEach(([agamaKey, tps]) => {
      tps.forEach((text, idx) => {
        toSave.push({
          id: `TP_${selectedKelas}_${makeSlug(mapelAgama)}_${makeSlug(agamaKey)}_TP${idx + 1}`,
          kodeTp: `TP ${idx + 1}`,
          mataPelajaran: mapelAgama,
          fase: selectedFase,
          kelas: selectedKelas,
          agama: agamaKey,
          tujuanPembelajaran: text
        });
      });
    });

    bulkSaveTujuanPembelajaran(toSave);
    setSelectedTpMapel(mapelAgama);
    showNotice(`Berhasil memuat ${toSave.length} TP 6 Agama resmi & menghubungkan ke rapor!`);
  };

  // Seed sample TPs & LMs for all empty subjects in current class
  const handleSeedAllSubjects = () => {
    const toSave: LearningObjective[] = [];
    db.mataPelajaran.forEach((m) => {
      const existing = db.tujuanPembelajaran.filter(
        (tp) => tp.kelas === selectedKelas && tp.mataPelajaran === m.mataPelajaran
      );
      if (existing.length === 0) {
        const samples = STANDARD_SAMPLE_TPS[m.mataPelajaran] || [
          `Siswa mampu memahami konsep inti pembelajaran ${m.mataPelajaran} dengan cermat.`,
          `Siswa mampu mendemonstrasikan keterampilan dan kreativitas pada mata pelajaran ${m.mataPelajaran}.`
        ];
        samples.forEach((text, idx) => {
          toSave.push({
            id: `TP_${selectedKelas}_${makeSlug(m.mataPelajaran)}_Semua_TP${idx + 1}`,
            kodeTp: `TP ${idx + 1}`,
            mataPelajaran: m.mataPelajaran,
            fase: selectedFase,
            kelas: selectedKelas,
            agama: 'Semua',
            tujuanPembelajaran: text
          });
        });
      }
      const lmSamples = STANDARD_SAMPLE_LMS[m.mataPelajaran];
      if (lmSamples) {
        const scopeKey = `${selectedKelas}::${m.mataPelajaran}`;
        const existingLm = { ...(db.lingkupMateriTitles?.[scopeKey] || {}) };
        let changed = false;
        lmSamples.forEach((name, idx) => {
          const k = `lm${idx + 1}`;
          if (!existingLm[k]) {
            existingLm[k] = name;
            changed = true;
          }
        });
        if (changed) saveLingkupMateriTitles(scopeKey, existingLm, true);
      }
    });

    if (toSave.length > 0) {
      bulkSaveTujuanPembelajaran(toSave);
      showNotice(`Berhasil memuat ${toSave.length} TP & Lingkup Materi untuk Kelas ${selectedKelas}.`);
    } else {
      showNotice('Semua mata pelajaran di kelas ini sudah memiliki TP.');
    }
  };

  // Export Excel
  const handleExportExcel = (targetMapel?: string) => {
    const mapelFilter = targetMapel || selectedTpMapel;
    const listToExport = db.tujuanPembelajaran.filter((tp) => {
      const matchesKelas = tp.kelas === selectedKelas;
      const matchesMapel = mapelFilter === 'SEMUA' || tp.mataPelajaran === mapelFilter;
      return matchesKelas && matchesMapel;
    });

    const wb = XLSX.utils.book_new();
    const data: (string | number)[][] = [
      [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
      [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
      [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
      ['LINGKUP MATERI DAN TUJUAN PEMBELAJARAN (TP) KURIKULUM MERDEKA'],
      [
        `Kelas: ${selectedKelas} • Fase: ${selectedFase} • Mata Pelajaran: ${
          mapelFilter === 'SEMUA' ? 'Semua Mata Pelajaran' : mapelFilter
        }`
      ],
      [],
      ['NO', 'KODE TP', 'MATA PELAJARAN', 'AGAMA', 'FASE', 'KELAS', 'RUMUSAN TUJUAN PEMBELAJARAN (TP)'],
      ['(1)', '(2)', '(3)', '(4)', '(5)', '(6)', '(7)'],
      ...listToExport.map((tp, idx) => [
        idx + 1,
        tp.kodeTp,
        tp.mataPelajaran,
        tp.agama || 'Semua',
        tp.fase,
        tp.kelas,
        tp.tujuanPembelajaran
      ])
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 12 },
      { wch: 34 },
      { wch: 14 },
      { wch: 8 },
      { wch: 10 },
      { wch: 65 }
    ];

    const sheetTitle = mapelFilter === 'SEMUA' ? 'Semua TP' : mapelFilter.slice(0, 30);
    XLSX.utils.book_append_sheet(wb, ws, sheetTitle);

    const filenameMapel =
      mapelFilter === 'SEMUA'
        ? 'semua-mapel'
        : mapelFilter
            .toLowerCase()
            .replace(/[^a-z0-9]/g, '-')
            .slice(0, 25);
    XLSX.writeFile(wb, `tp-${filenameMapel}-kelas-${selectedKelas}.xlsx`);
    showNotice(`File Excel TP (${mapelFilter}) berhasil diunduh.`);
  };

  // Import Excel
  const handleImportFile = async (file: File) => {
    try {
      const result = await parseAndImportTemplate(file, 'template_tp', appContext);
      showNotice(result.message);
    } catch (err: any) {
      console.error(err);
      showNotice('Gagal membaca file Excel TP: ' + (err?.message || 'Format berkas tidak sesuai'));
    }
  };

  // Download template
  const handleDownloadTemplate = (targetMapelName?: string) => {
    const mapel =
      targetMapelName && targetMapelName !== 'SEMUA'
        ? targetMapelName
        : selectedTpMapel !== 'SEMUA'
        ? selectedTpMapel
        : 'Semua Mata Pelajaran';
    const schoolName = db.sekolah.namaSekolah || 'SD NEGERI BALIL';
    const schoolYear = db.sekolah.tahunPelajaran || '2026/2027';
    const pemda = db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG';
    const instansi = db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN';

    const wb = XLSX.utils.book_new();
    const columns = [
      'Kode TP',
      'Mata Pelajaran',
      'Fase',
      'Kelas',
      'Rumusan Tujuan Pembelajaran (TP)',
      'Indikator Capaian / Lingkup Materi'
    ];

    const currentTps = db.tujuanPembelajaran.filter((t) => {
      const matchK = t.kelas === selectedKelas;
      const matchM = mapel === 'Semua Mata Pelajaran' || t.mataPelajaran === mapel;
      return matchK && matchM;
    });

    const sampleRows: (string | number)[][] = [];
    if (currentTps.length > 0) {
      currentTps.forEach((t) => {
        sampleRows.push([
          t.kodeTp,
          t.mataPelajaran,
          t.fase,
          t.kelas,
          t.tujuanPembelajaran,
          'Lingkup Materi'
        ]);
      });
    } else {
      sampleRows.push(
        [
          'TP 1',
          mapel === 'Semua Mata Pelajaran' ? 'Bahasa Indonesia' : mapel,
          selectedFase,
          selectedKelas,
          `Peserta didik mampu memahami konsep inti ${mapel}.`,
          'Lingkup Materi 1'
        ],
        [
          'TP 2',
          mapel === 'Semua Mata Pelajaran' ? 'Matematika' : mapel,
          selectedFase,
          selectedKelas,
          `Peserta didik mampu mempraktikkan keterampilan dasar pada ${mapel}.`,
          'Lingkup Materi 2'
        ]
      );
    }

    const aoaData: (string | number)[][] = [
      [pemda],
      [instansi],
      [schoolName],
      [`FORMAT RESMI TUJUAN PEMBELAJARAN (TP): ${mapel.toUpperCase()}`],
      [`Kelas: ${selectedKelas} • Fase: ${selectedFase} • Tahun Pelajaran: ${schoolYear}`],
      [],
      ['NO', ...columns],
      ['(1)', ...columns.map((_, i) => `(${i + 2})`)],
      ...sampleRows.map((r, i) => [i + 1, ...r])
    ];

    for (let i = 0; i < 12; i++) {
      aoaData.push([sampleRows.length + 1 + i, ...columns.map(() => '')]);
    }

    const ws = XLSX.utils.aoa_to_sheet(aoaData);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 12 },
      { wch: 32 },
      { wch: 10 },
      { wch: 10 },
      { wch: 55 },
      { wch: 30 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Template TP');
    const safeName = mapel.toLowerCase().replace(/[^a-z0-9]/g, '-');
    XLSX.writeFile(wb, `template-tp-${safeName}-kelas-${selectedKelas}.xlsx`);
    showNotice(`Template TP ${mapel} (.xlsx) berhasil diunduh!`);
  };

  // Filtered TPs for list_table mode
  const filteredTPs = useMemo(() => {
    return db.tujuanPembelajaran.filter((tp) => {
      const matchesKelas = tp.kelas === selectedKelas;
      const matchesMapel = selectedTpMapel === 'SEMUA' || tp.mataPelajaran === selectedTpMapel;
      const matchesAgama =
        selectedAgamaFilter === 'Semua' ||
        !tp.agama ||
        tp.agama === 'Semua' ||
        tp.agama.toLowerCase() === selectedAgamaFilter.toLowerCase();
      const matchesSearch =
        searchQuery.trim() === '' ||
        tp.kodeTp.toLowerCase().includes(searchQuery.toLowerCase()) ||
        tp.mataPelajaran.toLowerCase().includes(searchQuery.toLowerCase()) ||
        tp.tujuanPembelajaran.toLowerCase().includes(searchQuery.toLowerCase());
      return matchesKelas && matchesMapel && matchesAgama && matchesSearch;
    });
  }, [db.tujuanPembelajaran, selectedKelas, selectedTpMapel, selectedAgamaFilter, searchQuery]);

  const scrollToSubjectBlock = (mapelName: string) => {
    setSelectedTpMapel(mapelName);
    const elId = `sheet-mapel-${mapelName.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '')}`;
    const el = document.getElementById(elId);
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    }
  };

  const handleSyncFormulaNow = () => {
    const count = syncAllScoresWithTPDescriptions('ALL');
    showNotice(
      `Rumus TP dihubungkan ke Deskripsi Capaian Tertinggi, Terendah & Deskripsi Rapor (${count} data nilai diperbarui)!`
    );
  };

  return (
    <div className="space-y-3">
      {/* Hidden file input for Excel import */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) handleImportFile(f);
          e.target.value = '';
        }}
      />

      {/* Filter Nama Siswa Bar */}
      <div className="bg-white border border-slate-200 rounded-xl px-3.5 py-2 shadow-xs flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap items-center gap-2">
          <StudentFilterSelect />
          <div className="flex items-center gap-1.5 text-xs">
            <span className="font-bold text-slate-600">Kelas:</span>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value as ClassLevel)}
              aria-label="Pilih Kelas TP"
              className="px-2 py-1 rounded-lg border border-slate-300 text-xs font-extrabold text-slate-900 bg-slate-50"
            >
              {['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII'].map((k) => (
                <option key={k} value={k}>
                  Kelas {k}
                </option>
              ))}
            </select>
            <span className="px-2 py-0.5 rounded bg-cyan-100 text-cyan-950 font-bold text-[11px] border border-cyan-300">
              Fase {selectedFase}
            </span>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {notification && (
            <div className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg bg-emerald-800 text-white text-[11px] font-bold shadow-xs">
              <CheckCircle2 size={13} className="text-emerald-300 shrink-0" />
              <span className="max-w-[340px] truncate">{notification}</span>
            </div>
          )}

          {/* Switch Mode: Spreadsheet Excel (Default) vs Tabel Daftar */}
          <div className="flex items-center gap-0.5 bg-slate-100 p-0.5 rounded-lg border border-slate-200">
            <button
              type="button"
              onClick={() => setViewMode('spreadsheet')}
              title="Tampilan Lembar Kerja Excel (Lingkup Materi & TP 1–12)"
              aria-label="Mode Spreadsheet Excel"
              className={`w-7 h-7 inline-flex items-center justify-center rounded-md transition-colors cursor-pointer ${
                viewMode === 'spreadsheet'
                  ? 'bg-cyan-500 text-slate-950 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <LayoutGrid size={14} />
            </button>
            <button
              type="button"
              onClick={() => setViewMode('list_table')}
              title="Tampilan Tabel Daftar Lengkap"
              aria-label="Mode Tabel Daftar"
              className={`w-7 h-7 inline-flex items-center justify-center rounded-md transition-colors cursor-pointer ${
                viewMode === 'list_table'
                  ? 'bg-cyan-500 text-slate-950 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <TableIcon size={14} />
            </button>
          </div>
        </div>
      </div>

      {/* EXCEL SPREADSHEET WORKSPACE (MATCHING REFERENCE SCREENSHOT) */}
      <div className="rounded-xl border-2 border-slate-900 overflow-hidden shadow-md bg-cyan-300">
        <div className="flex items-stretch min-h-[560px]">
          {/* LEFT VERTICAL BLUE RAIL */}
          <aside className="w-[138px] shrink-0 bg-blue-700 border-r-2 border-slate-900 p-1.5 flex flex-col gap-1.5 select-none">
            {/* Home & Formula Info (?) Icon Row */}
            <div className="flex items-center justify-center gap-1.5 pb-1 border-b border-blue-500/70">
              <button
                type="button"
                onClick={() => setActiveMenu('dashboard')}
                title="Kembali ke Beranda / Dashboard"
                aria-label="Beranda"
                className="w-9 h-9 rounded-lg bg-sky-500 hover:bg-sky-400 text-white border border-white/60 shadow-xs inline-flex items-center justify-center cursor-pointer transition-colors"
              >
                <Home size={18} />
              </button>
              <button
                type="button"
                onClick={() => setShowFormulaHelp((v) => !v)}
                title="Info Rumus Koneksi TP ke Deskripsi Tertinggi, Terendah & Rapor"
                aria-label="Bantuan Rumus TP"
                className={`w-9 h-9 rounded-lg border shadow-xs inline-flex items-center justify-center cursor-pointer transition-colors ${
                  showFormulaHelp
                    ? 'bg-amber-400 text-slate-950 border-white'
                    : 'bg-pink-600 hover:bg-pink-500 text-white border-white/60'
                }`}
              >
                <HelpCircle size={18} />
              </button>
            </div>

            {/* 6 Agama Selector Buttons (Colorful Left Rail like Screenshot) */}
            <div className="space-y-1">
              <div className="text-[9px] font-black uppercase tracking-wider text-cyan-200 text-center">
                Filter Agama
              </div>
              {ENUM_6_AGAMA.map((ag) => {
                const isActive = selectedAgamaFilter === ag.key;
                return (
                  <button
                    key={ag.key}
                    type="button"
                    onClick={() => setSelectedAgamaFilter(ag.key)}
                    title={`Tampilkan & Edit TP Pendidikan Agama: ${ag.fullLabel}`}
                    className={`w-full py-1 px-1.5 rounded-sm border text-[10px] font-black tracking-tight text-center transition-all cursor-pointer ${
                      ag.btnClass
                    } ${isActive ? 'ring-2 ring-white scale-[1.02] shadow-sm' : 'opacity-90'}`}
                  >
                    {ag.label}
                  </button>
                );
              })}
            </div>

            {/* Quick Subject Jump Buttons */}
            <div className="mt-1 pt-1.5 border-t border-blue-500/70 space-y-1 flex-1 overflow-y-auto">
              <div className="text-[9px] font-black uppercase tracking-wider text-cyan-200 text-center">
                Lompat Mapel
              </div>
              {db.mataPelajaran.map((mapel, mIdx) => {
                const colorClass = SUBJECT_RAIL_COLORS[mIdx % SUBJECT_RAIL_COLORS.length];
                const isActive = selectedTpMapel === mapel.mataPelajaran;
                return (
                  <button
                    key={mapel.kode || mapel.mataPelajaran}
                    type="button"
                    onClick={() => scrollToSubjectBlock(mapel.mataPelajaran)}
                    title={mapel.mataPelajaran}
                    className={`w-full py-1 px-1.5 rounded-sm border border-slate-900 text-[9.5px] font-bold leading-tight text-center truncate cursor-pointer transition-all ${colorClass} ${
                      isActive ? 'ring-2 ring-white font-black' : ''
                    }`}
                  >
                    {mapel.mataPelajaran}
                  </button>
                );
              })}
            </div>
          </aside>

          {/* RIGHT MAIN CYAN WORKSHEET AREA */}
          <div className="flex-1 min-w-0 flex flex-col bg-cyan-300">
            {/* TOP CYAN BANNER + FORMULA INFO + ICON-ONLY TOOLBAR */}
            <div className="px-3 py-2 bg-cyan-300 border-b-2 border-slate-900 flex flex-wrap items-center justify-between gap-2">
              <div className="flex flex-wrap items-center gap-3">
                <div>
                  <h2 className="text-sm sm:text-base font-black tracking-wide text-slate-950 uppercase leading-tight">
                    LINGKUP MATERI DAN TUJUAN PEMBELAJARAN
                  </h2>
                  <p className="text-[10.5px] font-semibold text-slate-900">
                    Kelas {selectedKelas} (Fase {selectedFase}) · Agama Aktif:{' '}
                    <span className="font-black underline">{selectedAgamaFilter}</span> · Ketik langsung di sel putih — otomatis tersimpan &amp; terhubung ke Rapor.
                  </p>
                </div>

                {/* Formula box mirroring top-left box in Excel screenshot */}
                <div className="bg-white border border-slate-900 px-2.5 py-1 text-[10px] leading-snug text-slate-900 max-w-md shadow-2xs">
                  <span className="font-black text-blue-800">Rumus Otomatis Rapor:</span>{' '}
                  TP tertinggi &rarr; <strong>Deskripsi Capaian Tertinggi</strong> · TP terendah &rarr;{' '}
                  <strong>Deskripsi Capaian Terendah</strong> · Gabungan &rarr; <strong>Deskripsi Rapor</strong>.
                </div>
              </div>

              {/* ICON-ONLY COMPACT ACTION TOOLBAR */}
              <div className="flex flex-wrap items-center gap-1">
                <button
                  type="button"
                  onClick={handleSyncFormulaNow}
                  title="Hubungkan Rumus TP ke Deskripsi Capaian Tertinggi, Terendah & Deskripsi Rapor Sekarang"
                  aria-label="Hubungkan Rumus TP ke Deskripsi Rapor"
                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-blue-600 hover:bg-blue-700 text-white border border-slate-900 shadow-xs cursor-pointer"
                >
                  <Zap size={15} />
                </button>
                <button
                  type="button"
                  onClick={handleSeed6Agama}
                  title="Muat Contoh TP 6 Agama Resmi"
                  aria-label="Muat Contoh TP 6 Agama"
                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white border border-slate-900 shadow-xs cursor-pointer"
                >
                  <BookOpen size={15} />
                </button>
                <button
                  type="button"
                  onClick={handleSeedAllSubjects}
                  title="Isi Contoh Lingkup Materi & TP untuk Mapel yang Masih Kosong"
                  aria-label="Muat Contoh TP Semua Mapel"
                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-amber-400 hover:bg-amber-500 text-slate-950 border border-slate-900 shadow-xs cursor-pointer"
                >
                  <Sparkles size={15} />
                </button>
                <button
                  type="button"
                  onClick={() => handleOpenAdd()}
                  title="Tambah TP via Form Modal"
                  aria-label="Tambah TP"
                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-lime-400 hover:bg-lime-500 text-slate-950 border border-slate-900 shadow-xs cursor-pointer"
                >
                  <Plus size={16} />
                </button>
                <button
                  type="button"
                  onClick={() => fileInputRef.current?.click()}
                  title="Impor TP dari File Excel (.xlsx)"
                  aria-label="Impor Excel TP"
                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-white hover:bg-slate-100 text-slate-950 border border-slate-900 shadow-xs cursor-pointer"
                >
                  <Upload size={15} />
                </button>
                <button
                  type="button"
                  onClick={() => handleExportExcel()}
                  title="Expor Data TP ke Excel (.xlsx)"
                  aria-label="Expor Excel TP"
                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-white hover:bg-slate-100 text-slate-950 border border-slate-900 shadow-xs cursor-pointer"
                >
                  <Download size={15} />
                </button>
                <button
                  type="button"
                  onClick={() => handleDownloadTemplate(selectedTpMapel)}
                  title="Unduh Template Resmi TP (.xlsx)"
                  aria-label="Unduh Template TP"
                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-white hover:bg-slate-100 text-slate-950 border border-slate-900 shadow-xs cursor-pointer"
                >
                  <FileSpreadsheet size={15} />
                </button>
                <button
                  type="button"
                  onClick={() => setActiveMenu('input_nilai')}
                  title="Buka Input Nilai Siswa (Lihat Hasil Rumus TP di Deskripsi Rapor)"
                  aria-label="Buka Input Nilai Siswa"
                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg bg-pink-500 hover:bg-pink-600 text-white border border-slate-900 shadow-xs cursor-pointer"
                >
                  <Target size={15} />
                </button>
              </div>
            </div>

            {/* OPTIONAL COLLAPSIBLE FORMULA EXPLANATION PANEL */}
            {showFormulaHelp && (
              <div className="mx-3 mt-2 p-3 rounded-lg bg-white border-2 border-slate-900 text-xs text-slate-800 flex items-start justify-between gap-3 shadow-sm">
                <div className="space-y-1">
                  <div className="font-black text-blue-800 uppercase">
                    Cara Kerja Rumus Lingkup Materi &amp; Tujuan Pembelajaran (TP 1–TP 12):
                  </div>
                  <ul className="list-disc pl-4 space-y-0.5 text-[11px]">
                    <li>
                      Kolom <strong>Lingkup Materi (LM 1–LM 8)</strong> otomatis tampil sebagai judul kolom{' '}
                      <em>Sumatif Lingkup Materi</em> di menu <strong>Input Nilai Siswa</strong>.
                    </li>
                    <li>
                      Kolom <strong>TP 1–TP 12</strong> terhubung 1-ke-1 dengan kolom nilai <em>TP 1–TP 12</em> di menu{' '}
                      <strong>Input Nilai Siswa</strong>.
                    </li>
                    <li>
                      <strong>Deskripsi Capaian Tertinggi</strong> otomatis mengambil teks TP dengan nilai tertinggi (<code>&ge;85</code> sangat baik, <code>70–84</code> baik, <code>60–69</code> cukup).
                    </li>
                    <li>
                      <strong>Deskripsi Capaian Terendah</strong> otomatis mengambil teks TP dengan nilai terendah (<code>&lt;60</code> perlu pendampingan dan bimbingan intensif).
                    </li>
                    <li>
                      <strong>Deskripsi Rapor</strong> menggabungkan otomatis Capaian Tertinggi + Terendah dan langsung muncul di <strong>Cetak Rapor</strong> serta <strong>Buku Induk</strong>.
                    </li>
                  </ul>
                </div>
                <button
                  type="button"
                  onClick={() => setShowFormulaHelp(false)}
                  className="p-1 rounded hover:bg-slate-100 text-slate-600 cursor-pointer"
                  aria-label="Tutup bantuan"
                >
                  <X size={15} />
                </button>
              </div>
            )}

            {/* MODE 1: SPREADSHEET HORIZONTAL PER MATA PELAJARAN */}
            {viewMode === 'spreadsheet' && (
              <div className="p-3 overflow-x-auto flex-1">
                <div className="flex items-start gap-4 w-max pb-2">
                  {db.mataPelajaran.map((mapel) => {
                    const isAgamaSubject = mapel.mataPelajaran.toLowerCase().includes('agama');
                    const displaySubjectName =
                      isAgamaSubject && selectedAgamaFilter !== 'Semua'
                        ? `Pendidikan Agama ${selectedAgamaFilter} dan Budi Pekerti`
                        : mapel.mataPelajaran;

                    const matchedTps = getMatchedTpsForSubjectAndStudent(
                      mapel.mataPelajaran,
                      selectedKelas,
                      db.tujuanPembelajaran,
                      isAgamaSubject ? selectedAgamaFilter : undefined
                    );
                    const slots = resolveTpSlots(matchedTps);
                    const lmScopeKey = `${selectedKelas}::${mapel.mataPelajaran}`;
                    const lmTitles = db.lingkupMateriTitles?.[lmScopeKey] || {};
                    const blockId = `sheet-mapel-${mapel.mataPelajaran
                      .replace(/\s+/g, '-')
                      .replace(/[^a-zA-Z0-9-]/g, '')}`;
                    const effectiveAgamaKey = isAgamaSubject ? selectedAgamaFilter || 'Semua' : 'Semua';

                    return (
                      <div
                        key={mapel.kode || mapel.mataPelajaran}
                        id={blockId}
                        className="w-[455px] shrink-0 bg-cyan-300 select-text"
                      >
                        {/* Blue Subject Header */}
                        <div className="bg-blue-600 text-white border-2 border-slate-900 px-2 py-1.5 flex items-center justify-between gap-2">
                          <div className="w-5" />
                          <h3
                            className="text-xs font-black text-center tracking-wide truncate flex-1"
                            title={displaySubjectName}
                          >
                            {displaySubjectName}
                          </h3>
                          <div className="flex items-center gap-0.5 shrink-0">
                            <button
                              type="button"
                              onClick={() => handleSeedSamplesForMapel(mapel.mataPelajaran)}
                              title={`Isi Contoh LM & TP untuk ${displaySubjectName}`}
                              aria-label={`Isi contoh ${displaySubjectName}`}
                              className="w-5 h-5 inline-flex items-center justify-center rounded bg-blue-500 hover:bg-amber-400 hover:text-slate-950 text-white cursor-pointer transition-colors"
                            >
                              <Sparkles size={11} />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleExportExcel(mapel.mataPelajaran)}
                              title={`Expor Excel TP ${displaySubjectName}`}
                              aria-label={`Expor ${displaySubjectName}`}
                              className="w-5 h-5 inline-flex items-center justify-center rounded bg-blue-500 hover:bg-white hover:text-blue-900 text-white cursor-pointer transition-colors"
                            >
                              <Download size={11} />
                            </button>
                          </div>
                        </div>

                        {/* Lime Green 3-Column Sub-header */}
                        <div className="grid grid-cols-[128px_52px_1fr] bg-lime-400 text-slate-950 border-x-2 border-b-2 border-slate-900 font-black text-[11px] text-center">
                          <div className="py-1.5 px-1 border-r border-slate-900 flex items-center justify-center leading-tight">
                            Lingkup Materi
                          </div>
                          <div className="py-1.5 px-1 border-r border-slate-900 flex items-center justify-center">
                            TP
                          </div>
                          <div className="py-1.5 px-2 flex items-center justify-center">
                            Tujuan Pembelajaran (TP)
                          </div>
                        </div>

                        {/* 12 Rows: LM 1..8 on rows 1..8 (step cutout on rows 9..12), TP 1..12 on all 12 rows */}
                        <div>
                          {TP_COLUMN_KEYS.map((tpKey, rowIdx) => {
                            const hasLmSlot = rowIdx < LM_COLUMN_KEYS.length;
                            const lmKey = hasLmSlot ? LM_COLUMN_KEYS[rowIdx] : '';
                            const lmDraftKey = `${selectedKelas}::${mapel.mataPelajaran}::${lmKey}`;
                            const lmVal = hasLmSlot
                              ? lmDraftKey in lmDrafts
                                ? lmDrafts[lmDraftKey]
                                : lmTitles[lmKey] || ''
                              : '';

                            const slotTp = slots[rowIdx];
                            const tpDraftKey = `${selectedKelas}::${mapel.mataPelajaran}::${effectiveAgamaKey}::tp${rowIdx + 1}`;
                            const tpVal =
                              tpDraftKey in tpDrafts
                                ? tpDrafts[tpDraftKey]
                                : slotTp?.tujuanPembelajaran || '';

                            const isLastLmRow = rowIdx === LM_COLUMN_KEYS.length - 1;
                            const isLastTpRow = rowIdx === TP_COLUMN_KEYS.length - 1;

                            return (
                              <div
                                key={tpKey}
                                className="grid grid-cols-[128px_52px_1fr] items-stretch"
                              >
                                {/* Left column: LM 1..8 white input, rows 9..12 pure cyan cutout */}
                                {hasLmSlot ? (
                                  <div
                                    className={`bg-white border-l-2 border-r border-slate-900 ${
                                      isLastLmRow ? 'border-b-2' : 'border-b'
                                    } flex items-stretch`}
                                  >
                                    <input
                                      type="text"
                                      value={lmVal}
                                      onChange={(e) =>
                                        handleLmChange(mapel.mataPelajaran, lmKey, e.target.value)
                                      }
                                      onBlur={() => handleLmBlur(mapel.mataPelajaran, lmKey)}
                                      placeholder={`LM ${rowIdx + 1}...`}
                                      aria-label={`${mapel.mataPelajaran} Lingkup Materi ${rowIdx + 1}`}
                                      className="w-full px-1.5 py-1 text-[10.5px] font-semibold text-slate-950 bg-white placeholder-slate-300 focus:outline-hidden focus:bg-amber-50"
                                    />
                                  </div>
                                ) : (
                                  <div className="bg-cyan-300" />
                                )}

                                {/* Middle column: TP 1 .. TP 12 */}
                                <div
                                  className={`bg-white border-r border-slate-900 ${
                                    hasLmSlot ? '' : 'border-l-2'
                                  } ${
                                    isLastTpRow ? 'border-b-2' : 'border-b'
                                  } flex items-center justify-center text-[10.5px] font-extrabold text-slate-900 px-1 py-1`}
                                >
                                  {`TP ${rowIdx + 1}`}
                                </div>

                                {/* Right column: Tujuan Pembelajaran 1..12 inline editable */}
                                <div
                                  className={`bg-white border-r-2 border-slate-900 ${
                                    isLastTpRow ? 'border-b-2' : 'border-b'
                                  } flex items-stretch`}
                                >
                                  <textarea
                                    rows={2}
                                    value={tpVal}
                                    onChange={(e) =>
                                      handleTpChange(
                                        mapel.mataPelajaran,
                                        rowIdx,
                                        e.target.value,
                                        slotTp
                                      )
                                    }
                                    onBlur={() =>
                                      handleTpBlur(mapel.mataPelajaran, rowIdx, slotTp)
                                    }
                                    placeholder={`Isi rumusan TP ${rowIdx + 1}...`}
                                    aria-label={`${mapel.mataPelajaran} Tujuan Pembelajaran ${rowIdx + 1}`}
                                    className="w-full px-1.5 py-1 text-[10.5px] leading-snug text-slate-900 bg-white placeholder-slate-300 resize-none focus:outline-hidden focus:bg-amber-50"
                                  />
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* MODE 2: LIST TABLE VIEW (OPTIONAL ALTERNATIVE VIEW) */}
            {viewMode === 'list_table' && (
              <div className="p-3 space-y-3 bg-slate-50 flex-1">
                <div className="bg-white rounded-xl p-3 border border-slate-200 flex flex-wrap items-center justify-between gap-2">
                  <div className="flex flex-wrap items-center gap-2">
                    <select
                      value={selectedTpMapel}
                      onChange={(e) => setSelectedTpMapel(e.target.value)}
                      className="px-2.5 py-1.5 rounded-lg border border-slate-300 text-xs font-bold text-slate-800 bg-white"
                    >
                      <option value="SEMUA">Semua Mata Pelajaran</option>
                      {db.mataPelajaran.map((m) => (
                        <option key={m.kode || m.mataPelajaran} value={m.mataPelajaran}>
                          {m.mataPelajaran}
                        </option>
                      ))}
                    </select>
                    <span className="text-xs font-bold text-slate-600">
                      Menampilkan {filteredTPs.length} TP
                    </span>
                  </div>
                  <div className="relative w-full sm:w-64">
                    <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      placeholder="Cari rumusan atau kode TP..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full pl-8 pr-7 py-1.5 text-xs border border-slate-300 rounded-lg focus:outline-hidden focus:border-blue-600 bg-slate-50"
                    />
                    {searchQuery && (
                      <button
                        type="button"
                        onClick={() => setSearchQuery('')}
                        className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 cursor-pointer"
                      >
                        <X size={12} />
                      </button>
                    )}
                  </div>
                </div>

                <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-xs">
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs text-left">
                      <thead className="bg-slate-900 text-white uppercase text-[10px] tracking-wider font-bold">
                        <tr>
                          <th className="py-2.5 px-3 text-center w-10">No</th>
                          <th className="py-2.5 px-3 w-24">Kode TP</th>
                          <th className="py-2.5 px-3 w-48">Mata Pelajaran</th>
                          <th className="py-2.5 px-3 w-24 text-center">Agama</th>
                          <th className="py-2.5 px-3">Rumusan Tujuan Pembelajaran (TP)</th>
                          <th className="py-2.5 px-3 text-right w-24">Aksi</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredTPs.map((tp, idx) => (
                          <tr key={tp.id} className="hover:bg-slate-50">
                            <td className="py-2 px-3 text-center font-bold text-slate-400">{idx + 1}</td>
                            <td className="py-2 px-3 font-mono font-extrabold text-blue-700 whitespace-nowrap">
                              <span className="px-2 py-0.5 rounded bg-blue-50 border border-blue-200">
                                {tp.kodeTp}
                              </span>
                            </td>
                            <td className="py-2 px-3 font-semibold text-slate-800">
                              {tp.mataPelajaran}
                            </td>
                            <td className="py-2 px-3 text-center">
                              <span className="px-2 py-0.5 rounded bg-slate-100 text-slate-700 font-bold text-[10px]">
                                {tp.agama || 'Semua'}
                              </span>
                            </td>
                            <td className="py-2 px-3 text-slate-700 leading-relaxed">
                              {tp.tujuanPembelajaran}
                            </td>
                            <td className="py-2 px-3 text-right whitespace-nowrap">
                              <div className="inline-flex items-center gap-1">
                                <button
                                  type="button"
                                  onClick={() => handleDuplicateTP(tp)}
                                  title="Duplikat TP"
                                  className="p-1 hover:bg-slate-100 text-slate-600 rounded cursor-pointer"
                                >
                                  <Copy size={13} />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleOpenEdit(tp)}
                                  title="Edit TP"
                                  className="p-1 hover:bg-slate-100 text-blue-600 rounded cursor-pointer"
                                >
                                  <Edit2 size={13} />
                                </button>
                                {confirmDeleteId === tp.id ? (
                                  <button
                                    type="button"
                                    onClick={() => handleDelete(tp.id, tp.kodeTp)}
                                    title="Klik lagi untuk konfirmasi hapus"
                                    className="px-1.5 py-0.5 bg-rose-600 text-white rounded text-[10px] font-bold cursor-pointer"
                                  >
                                    Hapus?
                                  </button>
                                ) : (
                                  <button
                                    type="button"
                                    onClick={() => setConfirmDeleteId(tp.id)}
                                    title="Hapus TP"
                                    className="p-1 hover:bg-rose-50 text-rose-600 rounded cursor-pointer"
                                  >
                                    <Trash2 size={13} />
                                  </button>
                                )}
                              </div>
                            </td>
                          </tr>
                        ))}
                        {filteredTPs.length === 0 && (
                          <tr>
                            <td colSpan={6} className="py-10 text-center text-slate-500">
                              Belum ada data TP sesuai filter. Beralih ke tampilan Excel atau klik tombol + untuk menambah.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* MODAL FORM: TAMBAH / EDIT TUJUAN PEMBELAJARAN */}
      {modalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-4">
          <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full overflow-hidden border border-slate-200">
            <div className="bg-gradient-to-r from-blue-700 to-cyan-700 p-4 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-white/15 rounded-xl">
                  <Target size={18} />
                </div>
                <div>
                  <h3 className="font-bold text-sm">
                    {editingTP ? 'Edit Tujuan Pembelajaran' : 'Tambah Tujuan Pembelajaran Baru'}
                  </h3>
                  <p className="text-[11px] text-cyan-100">
                    Kurikulum Merdeka · Kelas {formData.kelas} (Fase {formData.fase})
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setModalOpen(false)}
                className="text-white/80 hover:text-white p-1 rounded-lg hover:bg-white/10 cursor-pointer"
              >
                <X size={18} />
              </button>
            </div>

            <form onSubmit={handleSubmit} className="p-5 space-y-3.5 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Kode Slot TP (TP 1 – TP 12) <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="text"
                    required
                    placeholder="Contoh: TP 1"
                    value={formData.kodeTp}
                    onChange={(e) => setFormData({ ...formData, kodeTp: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl font-mono font-bold text-blue-800 focus:outline-hidden focus:border-blue-600"
                  />
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Tingkat Kelas <span className="text-rose-500">*</span>
                  </label>
                  <select
                    value={formData.kelas}
                    onChange={(e) => {
                      const newKls = e.target.value as ClassLevel;
                      let newFase: FaseType = 'B';
                      if (newKls === 'I' || newKls === 'II') newFase = 'A';
                      else if (newKls === 'III' || newKls === 'IV') newFase = 'B';
                      else if (newKls === 'V' || newKls === 'VI') newFase = 'C';
                      else if (newKls === 'VII' || newKls === 'VIII' || newKls === 'IX') newFase = 'D';
                      else if (newKls === 'X') newFase = 'E';
                      else newFase = 'F';
                      setFormData({ ...formData, kelas: newKls, fase: newFase });
                    }}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl font-bold text-slate-800 focus:outline-hidden focus:border-blue-600"
                  >
                    {['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII'].map((k) => (
                      <option key={k} value={k}>
                        Kelas {k}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-slate-700 mb-1">
                    Mata Pelajaran <span className="text-rose-500">*</span>
                  </label>
                  <select
                    value={formData.mataPelajaran}
                    onChange={(e) => setFormData({ ...formData, mataPelajaran: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl font-bold text-slate-800 focus:outline-hidden focus:border-blue-600 bg-slate-50"
                  >
                    {db.mataPelajaran.map((m) => (
                      <option key={m.kode || m.mataPelajaran} value={m.mataPelajaran}>
                        {m.mataPelajaran}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block font-bold text-slate-700 mb-1">Kesesuaian Agama</label>
                  <select
                    value={formData.agama || 'Semua'}
                    onChange={(e) => setFormData({ ...formData, agama: e.target.value })}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl font-bold text-slate-800 focus:outline-hidden focus:border-blue-600 bg-slate-50"
                  >
                    <option value="Semua">Semua Agama (Umum)</option>
                    <option value="Islam">Islam (PAI)</option>
                    <option value="Kristen">Kristen Protestan (PAK)</option>
                    <option value="Katolik">Katolik (PAKAT)</option>
                    <option value="Hindu">Hindu (PAH)</option>
                    <option value="Buddha">Buddha (PAB)</option>
                    <option value="Khonghucu">Khonghucu (PAKh)</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 mb-1">
                  Rumusan Narasi Tujuan Pembelajaran (TP) <span className="text-rose-500">*</span>
                </label>
                <textarea
                  required
                  rows={4}
                  value={formData.tujuanPembelajaran}
                  onChange={(e) => setFormData({ ...formData, tujuanPembelajaran: e.target.value })}
                  placeholder="Contoh: Siswa mampu membaca, menulis, dan membandingkan bilangan cacah sampai 10.000."
                  className="w-full px-3 py-2 border border-slate-300 rounded-xl leading-relaxed focus:outline-hidden focus:border-blue-600 font-medium"
                />
                <div className="flex items-start gap-1.5 mt-1.5 p-2 rounded-lg bg-blue-50 border border-blue-200 text-[11px] text-blue-900">
                  <Info size={14} className="shrink-0 text-blue-600 mt-0.5" />
                  <span>
                    Rumusan TP ini otomatis terhubung ke <strong>Deskripsi Capaian Tertinggi</strong>,{' '}
                    <strong>Deskripsi Capaian Terendah</strong>, dan <strong>Deskripsi Rapor</strong> siswa.
                  </span>
                </div>
              </div>

              <div className="pt-3 border-t border-slate-200 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setModalOpen(false)}
                  className="px-4 py-2 border border-slate-300 rounded-xl font-bold text-slate-700 hover:bg-slate-100 cursor-pointer"
                >
                  Batal
                </button>
                <SaveStatusIndicator status={modalSaveStatus} />
                <button
                  type="submit"
                  disabled={modalSaveStatus === 'saving'}
                  className="px-5 py-2 bg-blue-700 hover:bg-blue-800 text-white rounded-xl font-bold flex items-center gap-1.5 shadow-sm cursor-pointer disabled:opacity-60"
                >
                  <Save size={14} /> Simpan TP
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
