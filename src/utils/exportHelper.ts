import * as XLSX from 'xlsx';
import { Student, AcademicScore, SchoolIdentity, ClassLevel, SemesterType, Subject } from '../types';
import { matchesStudentScore } from './studentSyncHelper';

/**
 * Utility helper to download CSV with UTF-8 BOM so Microsoft Excel
 * on Windows opens it with correct column separators and Indonesian characters.
 */
export function downloadCsv(filename: string, content: string, delimiter: ';' | ',' = ';') {
  const finalContent = `sep=${delimiter}\r\n` + content;
  const blob = new Blob(['\uFEFF' + finalContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', filename.endsWith('.csv') ? filename : `${filename}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * 1. EKSPOR DATA SISWA KE EXCEL (.xlsx)
 */
export function exportStudentsToExcel(
  students: Student[],
  school: SchoolIdentity,
  kelasFilter: string = 'SEMUA'
) {
  const wb = XLSX.utils.book_new();
  const titleKelas = kelasFilter === 'SEMUA' ? 'SEMUA KELAS' : `KELAS ${kelasFilter}`;

  const rows: (string | number)[][] = [
    [school.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
    [school.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
    [school.namaSekolah || 'SD NEGERI BALIL'],
    [`DATA POKOK PESERTA DIDIK - ${titleKelas}`],
    [`Tahun Pelajaran: ${school.tahunPelajaran} • Kurikulum Merdeka`],
    [],
    [
      'NO',
      'NISN',
      'NIS',
      'NAMA LENGKAP SISWA',
      'JK (L/P)',
      'TEMPAT LAHIR',
      'TANGGAL LAHIR',
      'AGAMA',
      'KELAS',
      'FASE',
      'SEMESTER',
      'NAMA AYAH',
      'NAMA IBU',
      'PEKERJAAN AYAH',
      'PEKERJAAN IBU',
      'NAMA WALI',
      'PEKERJAAN WALI',
      'ALAMAT WALI',
      'ALAMAT PESERTA DIDIK',
      'STATUS'
    ],
    ['(1)', '(2)', '(3)', '(4)', '(5)', '(6)', '(7)', '(8)', '(9)', '(10)', '(11)', '(12)', '(13)', '(14)', '(15)', '(16)', '(17)', '(18)', '(19)', '(20)'],
    ...students.map((s, idx) => [
      idx + 1,
      s.nisn,
      s.nis || '-',
      s.namaSiswa,
      s.jenisKelamin,
      s.tempatLahir || '-',
      s.tanggalLahir || '-',
      s.agama || 'Kristen',
      s.kelas,
      s.fase,
      s.semester || '1',
      s.namaAyah || '-',
      s.namaIbu || '-',
      s.pekerjaanAyah || '-',
      s.pekerjaanIbu || '-',
      s.namaWali || '-',
      s.pekerjaanWali || '-',
      s.alamatWali || '-',
      s.alamat || '-',
      s.statusSiswa || 'Aktif'
    ])
  ];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [
    { wch: 6 },
    { wch: 16 },
    { wch: 12 },
    { wch: 32 },
    { wch: 8 },
    { wch: 16 },
    { wch: 14 },
    { wch: 12 },
    { wch: 8 },
    { wch: 8 },
    { wch: 10 },
    { wch: 22 },
    { wch: 22 },
    { wch: 20 },
    { wch: 20 },
    { wch: 22 },
    { wch: 20 },
    { wch: 28 },
    { wch: 35 },
    { wch: 12 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Data Siswa');
  const safeFilename = `Data_Siswa_${school.namaSekolah.replace(/\s+/g, '_')}_Kelas_${kelasFilter}_TP_${school.tahunPelajaran.replace(/\//g, '-')}.xlsx`;
  XLSX.writeFile(wb, safeFilename);
}

/**
 * 1B. UNDUH TEMPLATE EXCEL DATA SISWA SIAP ISI & IMPOR (.xlsx)
 */
export function exportStudentTemplateToExcel(
  students: Student[],
  school: SchoolIdentity,
  kelasFilter: string = 'IV'
) {
  const wb = XLSX.utils.book_new();
  const targetKelas = kelasFilter === 'SEMUA' ? 'IV' : kelasFilter;

  const sampleOrActiveRows: (string | number)[][] =
    students.length > 0
      ? students.map((s, idx) => [
          idx + 1,
          s.nisn && !s.nisn.startsWith('BLM_') ? s.nisn : '',
          s.nis && s.nis !== '-' ? s.nis : '',
          s.namaSiswa,
          s.jenisKelamin,
          s.agama || 'Kristen',
          s.tempatLahir && s.tempatLahir !== '-' ? s.tempatLahir : '',
          s.tanggalLahir && s.tanggalLahir !== '-' ? s.tanggalLahir : '',
          s.nik && s.nik !== '-' ? s.nik : '',
          s.kelas || targetKelas,
          s.fase || 'B',
          s.namaAyah && s.namaAyah !== '-' ? s.namaAyah : '',
          s.namaIbu && s.namaIbu !== '-' ? s.namaIbu : '',
          s.pekerjaanAyah && s.pekerjaanAyah !== '-' && s.pekerjaanAyah !== '_' ? s.pekerjaanAyah : '',
          s.pekerjaanIbu && s.pekerjaanIbu !== '-' && s.pekerjaanIbu !== '_' ? s.pekerjaanIbu : '',
          s.namaWali && s.namaWali !== '-' && s.namaWali !== '_' ? s.namaWali : '',
          s.pekerjaanWali && s.pekerjaanWali !== '-' && s.pekerjaanWali !== '_' ? s.pekerjaanWali : '',
          s.alamatWali && s.alamatWali !== '-' && s.alamatWali !== '_' ? s.alamatWali : '',
          s.alamat || 'Kampung Balil, Oksibil',
          s.statusSiswa || 'Aktif'
        ])
      : [
          [1, '0098765432', '26001', 'Markus Nabyal', 'L', 'Kristen', 'Balil', '2016-04-12', '9124011204160001', targetKelas, 'B', 'Yohanes Nabyal', 'Maria Kasipmabin', 'Petani', 'Ibu Rumah Tangga', '', '', '', 'Kampung Balil, Oksibil', 'Aktif'],
          [2, '0098765433', '26002', 'Yohana Kasipmabin', 'P', 'Katolik', 'Oksibil', '2016-08-20', '9124012008160002', targetKelas, 'B', 'Anton Kasipmabin', 'Theresia Uropmabin', 'PNS', 'Ibu Rumah Tangga', 'Petrus Kasipmabin', 'Guru', 'Oksibil', 'Jalan Poros Balil', 'Aktif']
        ];

  const startEmptyNo = sampleOrActiveRows.length + 1;
  const blankRows: (string | number)[][] = Array.from({ length: 15 }, (_, i) => [
    startEmptyNo + i,
    '',
    '',
    '',
    'L',
    'Kristen',
    '',
    '',
    '',
    targetKelas,
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    'Aktif'
  ]);

  const rows: (string | number)[][] = [
    [school.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
    [school.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
    [school.namaSekolah || 'SD NEGERI BALIL'],
    ['TEMPLATE IMPOR DATA SISWA (PESERTA DIDIK) - KURIKULUM MERDEKA'],
    [`Tahun Pelajaran: ${school.tahunPelajaran} • Isi data mulai baris tabel di bawah ini, lalu simpan (.xlsx) dan impor kembali ke aplikasi.`],
    [],
    [
      'NO',
      'NISN',
      'NIS',
      'NAMA LENGKAP SISWA',
      'JK (L/P)',
      'AGAMA',
      'TEMPAT LAHIR',
      'TANGGAL LAHIR',
      'NIK',
      'KELAS',
      'FASE',
      'NAMA AYAH',
      'NAMA IBU',
      'PEKERJAAN AYAH',
      'PEKERJAAN IBU',
      'NAMA WALI',
      'PEKERJAAN WALI',
      'ALAMAT WALI',
      'ALAMAT PESERTA DIDIK',
      'STATUS'
    ],
    ['(1)', '(2)', '(3)', '(4)', '(5)', '(6)', '(7)', '(8)', '(9)', '(10)', '(11)', '(12)', '(13)', '(14)', '(15)', '(16)', '(17)', '(18)', '(19)', '(20)'],
    ...sampleOrActiveRows,
    ...blankRows
  ];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [
    { wch: 6 },
    { wch: 16 },
    { wch: 12 },
    { wch: 32 },
    { wch: 10 },
    { wch: 14 },
    { wch: 16 },
    { wch: 15 },
    { wch: 18 },
    { wch: 9 },
    { wch: 8 },
    { wch: 24 },
    { wch: 24 },
    { wch: 20 },
    { wch: 20 },
    { wch: 22 },
    { wch: 20 },
    { wch: 28 },
    { wch: 36 },
    { wch: 12 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Template Data Siswa');
  XLSX.writeFile(wb, `template-data-siswa-kelas-${kelasFilter}.xlsx`);
}

/**
 * 2. EKSPOR DATA SISWA KE CSV (.csv)
 */
export function exportStudentsToCsv(
  students: Student[],
  school: SchoolIdentity,
  kelasFilter: string = 'SEMUA'
) {
  const headers = [
    'No',
    'NISN',
    'NIS',
    'Nama Siswa',
    'JK',
    'Tempat Lahir',
    'Tanggal Lahir',
    'Agama',
    'Kelas',
    'Fase',
    'Semester',
    'Nama Ayah',
    'Nama Ibu',
    'Pekerjaan Ayah',
    'Pekerjaan Ibu',
    'Nama Wali',
    'Pekerjaan Wali',
    'Alamat Wali',
    'Alamat',
    'Status'
  ];

  const headerRow = headers.map((h) => `"${h}"`).join(';');
  const dataRows = students.map((s, idx) => {
    return [
      idx + 1,
      `"${s.nisn}"`,
      `"${s.nis || '-'}"`,
      `"${s.namaSiswa.replace(/"/g, '""')}"`,
      `"${s.jenisKelamin}"`,
      `"${(s.tempatLahir || '-').replace(/"/g, '""')}"`,
      `"${s.tanggalLahir || '-'}"`,
      `"${s.agama || 'Kristen'}"`,
      `"${s.kelas}"`,
      `"${s.fase}"`,
      `"${s.semester || '1'}"`,
      `"${(s.namaAyah || '-').replace(/"/g, '""')}"`,
      `"${(s.namaIbu || '-').replace(/"/g, '""')}"`,
      `"${(s.pekerjaanAyah || '-').replace(/"/g, '""')}"`,
      `"${(s.pekerjaanIbu || '-').replace(/"/g, '""')}"`,
      `"${(s.namaWali || '-').replace(/"/g, '""')}"`,
      `"${(s.pekerjaanWali || '-').replace(/"/g, '""')}"`,
      `"${(s.alamatWali || '-').replace(/"/g, '""')}"`,
      `"${(s.alamat || '-').replace(/"/g, '""')}"`,
      `"${s.statusSiswa || 'Aktif'}"`
    ].join(';');
  });

  const csvString = [headerRow, ...dataRows].join('\r\n');
  const safeFilename = `Data_Siswa_Kelas_${kelasFilter}_TP_${school.tahunPelajaran.replace(/\//g, '-')}.csv`;
  downloadCsv(safeFilename, csvString, ';');
}

/**
 * 3. EKSPOR NILAI MATA PELAJARAN AKTIF KE EXCEL (.xlsx)
 */
export function exportSubjectScoresToExcel(
  scores: AcademicScore[],
  mapel: string,
  kelas: ClassLevel | string,
  semester: SemesterType,
  school: SchoolIdentity
) {
  const wb = XLSX.utils.book_new();

  const rows: (string | number)[][] = [
    [school.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
    [school.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
    [school.namaSekolah || 'SD NEGERI BALIL'],
    [`REKAP NILAI KURIKULUM MERDEKA - MATA PELAJARAN: ${mapel.toUpperCase()}`],
    [`Kelas: ${kelas} • Semester: ${semester} • Tahun Pelajaran: ${school.tahunPelajaran}`],
    [],
    [
      'NO',
      'NISN',
      'NAMA SISWA',
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
      'NON TES',
      'TES',
      'NA SUMATIF TP',
      'NA SUMATIF LM',
      'NA SUMATIF SAS',
      'NILAI AKHIR (RAPOR)',
      'DESKRIPSI CAPAIAN TERTINGGI',
      'DESKRIPSI CAPAIAN TERENDAH'
    ],
    ...scores.map((sc, idx) => [
      idx + 1,
      sc.nisn,
      sc.namaSiswa,
      sc.tpScores?.tp1 ?? '',
      sc.tpScores?.tp2 ?? '',
      sc.tpScores?.tp3 ?? '',
      sc.tpScores?.tp4 ?? '',
      sc.tpScores?.tp5 ?? '',
      sc.tpScores?.tp6 ?? '',
      sc.tpScores?.tp7 ?? '',
      sc.tpScores?.tp8 ?? '',
      sc.lmScores?.lm1 ?? '',
      sc.lmScores?.lm2 ?? '',
      sc.lmScores?.lm3 ?? '',
      sc.lmScores?.lm4 ?? '',
      sc.sasNonTes ?? '',
      sc.sasTes ?? '',
      sc.naSumatifTp !== undefined ? sc.naSumatifTp.toFixed(2) : '',
      sc.naSumatifLm !== undefined ? sc.naSumatifLm.toFixed(2) : '',
      sc.naSumatifSas !== undefined ? sc.naSumatifSas.toFixed(2) : '',
      sc.nilaiAkhir || '',
      sc.deskripsiTercapai || '',
      sc.deskripsiPerluPeningkatan || ''
    ])
  ];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [
    { wch: 5 },
    { wch: 16 },
    { wch: 28 },
    { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 },
    { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 },
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 },
    { wch: 9 }, { wch: 9 },
    { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
    { wch: 45 }, { wch: 45 }
  ];

  const sheetName = `Nilai ${mapel.replace(/[:\\/?*[\]]/g, '').slice(0, 25)}`;
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  const safeFilename = `Nilai_${mapel.replace(/\s+/g, '_')}_Kelas_${kelas}_Sem_${semester}.xlsx`;
  XLSX.writeFile(wb, safeFilename);
}

/**
 * 4. EKSPOR NILAI MATA PELAJARAN AKTIF KE CSV (.csv)
 */
export function exportSubjectScoresToCsv(
  scores: AcademicScore[],
  mapel: string,
  kelas: ClassLevel | string,
  semester: SemesterType,
  school: SchoolIdentity
) {
  void school;
  const headers = [
    'No',
    'NISN',
    'Nama Siswa',
    'Mata Pelajaran',
    'Kelas',
    'Semester',
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

  const headerRow = headers.map((h) => `"${h}"`).join(';');
  const dataRows = scores.map((sc, idx) => {
    return [
      idx + 1,
      `"${sc.nisn}"`,
      `"${sc.namaSiswa.replace(/"/g, '""')}"`,
      `"${mapel.replace(/"/g, '""')}"`,
      `"${kelas}"`,
      `"${semester}"`,
      sc.tpScores?.tp1 ?? '',
      sc.tpScores?.tp2 ?? '',
      sc.tpScores?.tp3 ?? '',
      sc.tpScores?.tp4 ?? '',
      sc.tpScores?.tp5 ?? '',
      sc.tpScores?.tp6 ?? '',
      sc.tpScores?.tp7 ?? '',
      sc.tpScores?.tp8 ?? '',
      sc.lmScores?.lm1 ?? '',
      sc.lmScores?.lm2 ?? '',
      sc.lmScores?.lm3 ?? '',
      sc.lmScores?.lm4 ?? '',
      sc.sasNonTes ?? '',
      sc.sasTes ?? '',
      sc.naSumatifTp !== undefined ? sc.naSumatifTp.toFixed(2) : '',
      sc.naSumatifLm !== undefined ? sc.naSumatifLm.toFixed(2) : '',
      sc.naSumatifSas !== undefined ? sc.naSumatifSas.toFixed(2) : '',
      sc.nilaiAkhir || '',
      `"${(sc.deskripsiTercapai || '').replace(/"/g, '""')}"`,
      `"${(sc.deskripsiPerluPeningkatan || '').replace(/"/g, '""')}"`
    ].join(';');
  });

  const csvString = [headerRow, ...dataRows].join('\r\n');
  const safeFilename = `Nilai_${mapel.replace(/\s+/g, '_')}_Kelas_${kelas}_Sem_${semester}.csv`;
  downloadCsv(safeFilename, csvString, ';');
}

/**
 * 5. EKSPOR BUKU NILAI KOMPLIT WALI KELAS (MULTI-SHEET WORKBOOK .XLSX)
 * Memuat:
 * Sheet 1: Leger Nilai & Peringkat Kelas
 * Sheet 2: Data Pokok Siswa
 * Sheet 3+: Lembar Nilai Masing-Masing Mata Pelajaran
 */
export function exportClassComprehensiveWorkbook(
  db: {
    sekolah: SchoolIdentity;
    siswa: Student[];
    mataPelajaran: Subject[];
    nilaiAkademik: AcademicScore[];
  },
  kelas: ClassLevel | string,
  semester: SemesterType
) {
  const wb = XLSX.utils.book_new();
  const siswaKelas = db.siswa.filter((s) => s.kelas === kelas);
  const mapelList = db.mataPelajaran;

  // --- SHEET 1: LEGER NILAI ---
  const matrix = siswaKelas.map((s) => {
    const scores = mapelList.map((m) => {
      const match = db.nilaiAkademik.find(
        (n) => matchesStudentScore(n, s) && n.mataPelajaran === m.mataPelajaran && n.semester === semester
      );
      return match && match.nilaiAkhir > 0 ? match.nilaiAkhir : null;
    });

    const validScores = scores.filter((v): v is number => v !== null);
    const total = validScores.reduce((acc, curr) => acc + curr, 0);
    const avg = validScores.length > 0 ? parseFloat((total / validScores.length).toFixed(1)) : 0;

    return {
      student: s,
      scores,
      total,
      avg
    };
  });

  const ranked = [...matrix].sort((a, b) => b.avg - a.avg);
  const rankMap = new Map<string, number>();
  ranked.forEach((item, idx) => {
    rankMap.set(item.student.idSiswa, idx + 1);
  });

  const legerRows: (string | number)[][] = [
    [db.sekolah.namaPemerintah || 'PEMERINTAH KABUPATEN PEGUNUNGAN BINTANG'],
    [db.sekolah.namaInstansi || 'DINAS PENDIDIKAN DAN PENGAJARAN'],
    [db.sekolah.namaSekolah || 'SD NEGERI BALIL'],
    [`BUKU LEGER NILAI & PERINGKAT KELAS ${kelas}`],
    [`Semester: ${semester} • Tahun Pelajaran: ${db.sekolah.tahunPelajaran} • Kurikulum Merdeka`],
    [],
    [
      'RANK',
      'NISN',
      'NAMA SISWA',
      'JK',
      ...mapelList.map((m) => m.mataPelajaran),
      'TOTAL NILAI',
      'RATA-RATA'
    ],
    [
      '(1)',
      '(2)',
      '(3)',
      '(4)',
      ...mapelList.map((_, i) => `(${i + 5})`),
      `(${mapelList.length + 5})`,
      `(${mapelList.length + 6})`
    ],
    ...ranked.map((item) => [
      rankMap.get(item.student.idSiswa) || '-',
      item.student.nisn,
      item.student.namaSiswa,
      item.student.jenisKelamin,
      ...item.scores.map((sc) => sc ?? '-'),
      item.total,
      item.avg
    ])
  ];

  const wsLeger = XLSX.utils.aoa_to_sheet(legerRows);
  wsLeger['!cols'] = [
    { wch: 8 },
    { wch: 16 },
    { wch: 28 },
    { wch: 6 },
    ...mapelList.map(() => ({ wch: 14 })),
    { wch: 14 },
    { wch: 14 }
  ];
  XLSX.utils.book_append_sheet(wb, wsLeger, 'Leger Nilai');

  // --- SHEET 2: DATA SISWA ---
  const studentRows: (string | number)[][] = [
    ['NO', 'NISN', 'NIS', 'NAMA SISWA', 'JK', 'TEMPAT LAHIR', 'TGL LAHIR', 'AGAMA', 'KELAS', 'FASE', 'NAMA AYAH', 'NAMA IBU', 'PEKERJAAN AYAH', 'PEKERJAAN IBU', 'NAMA WALI', 'PEKERJAAN WALI', 'ALAMAT WALI', 'ALAMAT'],
    ...siswaKelas.map((s, idx) => [
      idx + 1,
      s.nisn,
      s.nis || '-',
      s.namaSiswa,
      s.jenisKelamin,
      s.tempatLahir || '-',
      s.tanggalLahir || '-',
      s.agama || 'Kristen',
      s.kelas,
      s.fase,
      s.namaAyah || '-',
      s.namaIbu || '-',
      s.pekerjaanAyah || '-',
      s.pekerjaanIbu || '-',
      s.namaWali || '-',
      s.pekerjaanWali || '-',
      s.alamatWali || '-',
      s.alamat || '-'
    ])
  ];
  const wsStudents = XLSX.utils.aoa_to_sheet(studentRows);
  XLSX.utils.book_append_sheet(wb, wsStudents, 'Data Siswa');

  // --- SHEET 3+: NILAI PER MATA PELAJARAN ---
  mapelList.forEach((mapel) => {
    const mapelScores = siswaKelas.map((s) => {
      const found = db.nilaiAkademik.find(
        (n) => matchesStudentScore(n, s) && n.mataPelajaran === mapel.mataPelajaran && n.semester === semester
      );
      return found || {
        nisn: s.nisn,
        namaSiswa: s.namaSiswa,
        formatif: undefined,
        sumatif: undefined,
        nilaiAkhir: 0,
        predikat: 'Perlu Bimbingan',
        capaianKompetensi: 'Perlu Pendampingan',
        deskripsiRapor: '-'
      };
    });

    const mapelRows: (string | number)[][] = [
      [`NILAI: ${mapel.mataPelajaran.toUpperCase()} - KELAS ${kelas} SEMESTER ${semester}`],
      [],
      ['NO', 'NISN', 'NAMA SISWA', 'FORMATIF', 'SUMATIF', 'NILAI AKHIR', 'PREDIKAT', 'CAPAIAN', 'DESKRIPSI RAPOR'],
      ...mapelScores.map((sc, idx) => [
        idx + 1,
        sc.nisn,
        sc.namaSiswa,
        sc.formatif ?? '',
        sc.sumatif ?? '',
        sc.nilaiAkhir,
        sc.predikat,
        sc.capaianKompetensi,
        sc.deskripsiRapor
      ])
    ];

    const wsMapel = XLSX.utils.aoa_to_sheet(mapelRows);
    wsMapel['!cols'] = [
      { wch: 6 },
      { wch: 16 },
      { wch: 28 },
      { wch: 10 },
      { wch: 10 },
      { wch: 12 },
      { wch: 14 },
      { wch: 16 },
      { wch: 45 }
    ];
    const cleanSheetName = mapel.mataPelajaran.replace(/[:\\/?*[\]]/g, '').slice(0, 25);
    XLSX.utils.book_append_sheet(wb, wsMapel, cleanSheetName);
  });

  const safeFilename = `Buku_Nilai_Komplit_Wali_Kelas_${kelas}_Sem_${semester}_${db.sekolah.namaSekolah.replace(/\s+/g, '_')}.xlsx`;
  XLSX.writeFile(wb, safeFilename);
}

/**
 * 6. EKSPOR LEGER NILAI SEMUA MATA PELAJARAN KE CSV (.csv)
 */
export function exportClassLegerToCsv(
  db: {
    sekolah: SchoolIdentity;
    siswa: Student[];
    mataPelajaran: Subject[];
    nilaiAkademik: AcademicScore[];
  },
  kelas: ClassLevel | string,
  semester: SemesterType
) {
  const siswaKelas = db.siswa.filter((s) => s.kelas === kelas);
  const mapelList = db.mataPelajaran;

  const matrix = siswaKelas.map((s) => {
    const scores = mapelList.map((m) => {
      const match = db.nilaiAkademik.find(
        (n) => matchesStudentScore(n, s) && n.mataPelajaran === m.mataPelajaran && n.semester === semester
      );
      return match ? match.nilaiAkhir : null;
    });

    const validScores = scores.filter((v): v is number => v !== null);
    const total = validScores.reduce((acc, curr) => acc + curr, 0);
    const avg = validScores.length > 0 ? parseFloat((total / validScores.length).toFixed(1)) : 0;

    return {
      student: s,
      scores,
      total,
      avg
    };
  });

  const ranked = [...matrix].sort((a, b) => b.avg - a.avg);
  const rankMap = new Map<string, number>();
  ranked.forEach((item, idx) => {
    rankMap.set(item.student.idSiswa, idx + 1);
  });

  const headerCols = [
    'Peringkat',
    'NISN',
    'Nama Siswa',
    'JK',
    ...mapelList.map((m) => m.mataPelajaran),
    'Total Nilai',
    'Rata-rata'
  ];

  const headerRow = headerCols.map((h) => `"${h}"`).join(';');
  const dataRows = ranked.map((item) => {
    const rank = rankMap.get(item.student.idSiswa) || '-';
    const scoreCols = item.scores.map((sc) => sc ?? '-').join(';');
    return `"${rank}";"${item.student.nisn}";"${item.student.namaSiswa.replace(/"/g, '""')}";"${item.student.jenisKelamin}";${scoreCols};${item.total};${item.avg}`;
  });

  const csvString = [headerRow, ...dataRows].join('\r\n');
  const safeFilename = `Leger_Nilai_Kelas_${kelas}_Sem_${semester}.csv`;
  downloadCsv(safeFilename, csvString, ';');
}
