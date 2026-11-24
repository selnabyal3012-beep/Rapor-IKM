import {
  DatabaseState,
  Student,
  ClassData,
  ClassLevel,
  SemesterType,
  AcademicScore,
  Attendance,
  TeacherNote,
  Extracurricular,
  StudentDecision
} from '../types';
import {
  generateReportDescriptionFromTP,
  getMatchedTpsForSubjectAndStudent,
  calculateKurikulumMerdekaScores,
  hasKurikulumMerdekaInputs,
  calculateScoreSummaryFromFreeInputs,
  getPredikatFromScore,
  getCapaianFromScore,
  formatStudentCallName,
  SCORE_INPUT_FIELDS
} from './tpDescriptionHelper';

export function isPlaceholderNisn(nisn?: string): boolean {
  const clean = (nisn || '').trim().toLowerCase();
  return !clean || clean === '-' || clean === '_' || clean === 'kosong' || clean === 'belum ada' || clean === 'null' || clean === 'undefined';
}

export function matchesStudentScore(
  score: { idSiswa?: string; nisn?: string; namaSiswa?: string; kelas?: string },
  student: Pick<Student, 'idSiswa' | 'nisn' | 'namaSiswa'> & { kelas?: string }
): boolean {
  if (score.idSiswa && student.idSiswa) {
    return score.idSiswa === student.idSiswa;
  }
  if (!isPlaceholderNisn(score.nisn) && !isPlaceholderNisn(student.nisn)) {
    return score.nisn === student.nisn;
  }
  if (score.namaSiswa && student.namaSiswa) {
    const sameName = score.namaSiswa.trim().toLowerCase() === student.namaSiswa.trim().toLowerCase();
    if (!sameName) return false;
    if (score.kelas && student.kelas) {
      return score.kelas === student.kelas;
    }
    return true;
  }
  return false;
}

export function matchesStudentRecord(
  rec: { idSiswa?: string; nisn?: string; namaSiswa?: string; kelas?: string },
  student: Pick<Student, 'idSiswa' | 'nisn' | 'namaSiswa'> & { kelas?: string }
): boolean {
  if (rec.idSiswa && student.idSiswa) {
    return rec.idSiswa === student.idSiswa;
  }
  if (!isPlaceholderNisn(rec.nisn) && !isPlaceholderNisn(student.nisn)) {
    return rec.nisn === student.nisn;
  }
  if (rec.namaSiswa && student.namaSiswa) {
    const sameName = rec.namaSiswa.trim().toLowerCase() === student.namaSiswa.trim().toLowerCase();
    if (!sameName) return false;
    if (rec.kelas && student.kelas) {
      return rec.kelas === student.kelas;
    }
    return true;
  }
  return false;
}

/**
 * Ensures all students in DatabaseState have non-empty, unique idSiswa and nisn values,
 * and backfills idSiswa + canonical nisn + namaSiswa across every student-related collection.
 */
export function ensureUniqueStudentNisns(prevDb: DatabaseState): DatabaseState {
  const siswaRaw = prevDb.siswa || [];
  const usedIds = new Set<string>();
  const usedNisns = new Set<string>();
  let placeholderCount = 0;

  siswaRaw.forEach((s) => {
    const key = (s.nisn || '').trim();
    if (isPlaceholderNisn(key)) {
      placeholderCount++;
    }
  });

  let changed = false;
  const fixedSiswa = siswaRaw.map((s, idx) => {
    let idSiswa = (s.idSiswa || '').trim();
    if (!idSiswa || usedIds.has(idSiswa)) {
      idSiswa = `S_AUTO_${idx + 1}_${Date.now().toString(36).slice(-4)}`;
      changed = true;
    }
    usedIds.add(idSiswa);

    let nisn = (s.nisn || '').trim();
    if (isPlaceholderNisn(nisn) || usedNisns.has(nisn)) {
      nisn = `BLM_${idSiswa}`;
      while (usedNisns.has(nisn)) {
        nisn = `BLM_${idSiswa}_${idx + 1}`;
      }
      changed = true;
    }
    usedNisns.add(nisn);

    if (idSiswa !== s.idSiswa || nisn !== s.nisn) {
      return { ...s, idSiswa, nisn };
    }
    return s;
  });

  const singlePlaceholderStudent =
    placeholderCount === 1
      ? fixedSiswa.find((_, i) => isPlaceholderNisn(siswaRaw[i]?.nisn))
      : undefined;

  const findUniqueStudentByName = (nama?: string, kelas?: string): Student | undefined => {
    const cleanName = (nama || '').trim().toLowerCase();
    if (!cleanName) return undefined;
    const matchesByClass = kelas
      ? fixedSiswa.filter(
          (s) => s.namaSiswa.trim().toLowerCase() === cleanName && s.kelas === kelas
        )
      : [];
    if (matchesByClass.length === 1) return matchesByClass[0];
    const allMatches = fixedSiswa.filter((s) => s.namaSiswa.trim().toLowerCase() === cleanName);
    return allMatches.length === 1 ? allMatches[0] : undefined;
  };

  const resolveLinkedStudent = (rec: {
    idSiswa?: string;
    nisn?: string;
    namaSiswa?: string;
    kelas?: string;
  }): { matched?: Student; drop: boolean } => {
    if (rec.idSiswa) {
      const byId = fixedSiswa.find((s) => s.idSiswa === rec.idSiswa);
      if (byId) return { matched: byId, drop: false };
    }
    if (!isPlaceholderNisn(rec.nisn)) {
      const byNisn = fixedSiswa.find((s) => s.nisn === rec.nisn);
      if (byNisn) return { matched: byNisn, drop: false };
    }
    const byName = findUniqueStudentByName(rec.namaSiswa, rec.kelas);
    if (byName) return { matched: byName, drop: false };
    if (isPlaceholderNisn(rec.nisn)) {
      if (singlePlaceholderStudent) {
        return { matched: singlePlaceholderStudent, drop: false };
      }
      return { drop: true };
    }
    return { drop: false };
  };

  let collectionsChanged = false;

  const fixedNilai = (prevDb.nilaiAkademik || []).reduce<AcademicScore[]>((acc, n) => {
    const { matched, drop } = resolveLinkedStudent(n);
    if (drop) {
      collectionsChanged = true;
      return acc;
    }
    if (matched) {
      if (n.idSiswa !== matched.idSiswa || n.nisn !== matched.nisn || n.namaSiswa !== matched.namaSiswa) {
        collectionsChanged = true;
        acc.push({ ...n, idSiswa: matched.idSiswa, nisn: matched.nisn, namaSiswa: matched.namaSiswa });
        return acc;
      }
    }
    acc.push(n);
    return acc;
  }, []);

  const remapCollection = <T extends { idSiswa?: string; nisn: string; namaSiswa?: string; kelas?: string }>(
    list: T[] | undefined
  ): T[] =>
    (list || []).reduce<T[]>((acc, item) => {
      const { matched, drop } = resolveLinkedStudent(item);
      if (drop) {
        collectionsChanged = true;
        return acc;
      }
      if (matched) {
        const nextNama = item.namaSiswa !== undefined ? matched.namaSiswa : item.namaSiswa;
        if (
          item.idSiswa !== matched.idSiswa ||
          item.nisn !== matched.nisn ||
          (item.namaSiswa !== undefined && item.namaSiswa !== matched.namaSiswa)
        ) {
          collectionsChanged = true;
          acc.push({
            ...item,
            idSiswa: matched.idSiswa,
            nisn: matched.nisn,
            ...(item.namaSiswa !== undefined ? { namaSiswa: nextNama } : {})
          });
          return acc;
        }
      }
      acc.push(item);
      return acc;
    }, []);

  const fixedKehadiran = remapCollection(prevDb.kehadiran);
  const fixedCatatan = remapCollection(prevDb.catatanWali);
  const fixedEkskul = remapCollection(prevDb.ekstrakurikuler);
  const fixedKeputusan = remapCollection(prevDb.keputusanSiswa);
  const fixedPiagam = remapCollection(prevDb.piagam);
  const fixedP5Nilai = remapCollection(prevDb.p5Nilai);

  if (!changed && !collectionsChanged) {
    return prevDb;
  }

  return {
    ...prevDb,
    siswa: fixedSiswa,
    nilaiAkademik: fixedNilai,
    kehadiran: fixedKehadiran,
    catatanWali: fixedCatatan,
    ekstrakurikuler: fixedEkskul,
    keputusanSiswa: fixedKeputusan,
    piagam: fixedPiagam,
    p5Nilai: fixedP5Nilai
  };
}

/**
 * Generates a collision-free unique ID for P5Score based on project, NISN, and full dimension name
 */
export function buildUniqueP5ScoreId(idProjek: string, nisn: string, dimensi: string): string {
  const slug =
    (dimensi || 'umum')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '') || 'umum';
  return `NP_${idProjek}_${nisn}_${slug}`;
}

/**
 * Returns the next grade level according to Indonesian curriculum standard
 */
export function getNextClass(current: ClassLevel): string {
  switch (current) {
    case 'I':
      return 'II (Dua)';
    case 'II':
      return 'III (Tiga)';
    case 'III':
      return 'IV (Empat)';
    case 'IV':
      return 'V (Lima)';
    case 'V':
      return 'VI (Enam)';
    case 'VI':
      return 'Lulus dari Satuan Pendidikan SD';
    case 'VII':
      return 'VIII (Delapan)';
    case 'VIII':
      return 'IX (Sembilan)';
    case 'IX':
      return 'Lulus dari Satuan Pendidikan SMP';
    case 'X':
      return 'XI (Sebelas)';
    case 'XI':
      return 'XII (Dua Belas)';
    case 'XII':
      return 'Lulus dari Satuan Pendidikan SMA/SMK';
    case 'XIII':
      return 'Lulus dari Satuan Pendidikan SMK Program 4 Tahun';
    default:
      return 'Tingkat Berikutnya';
  }
}

/**
 * Automatically adjusts the subject name for religion according to the student's religion
 * e.g. "Pendidikan Agama dan Budi Pekerti" + "Katolik" -> "Pendidikan Agama Katolik dan Budi Pekerti"
 */
export function getNamaPelajaranAgama(baseMapelName: string, agama?: string): string {
  if (!baseMapelName) return '';
  const lowerBase = baseMapelName.toLowerCase();
  if (!lowerBase.includes('agama')) {
    return baseMapelName;
  }
  const cleanAgama = (agama || '').trim();
  if (!cleanAgama) return baseMapelName;

  const lowerAgama = cleanAgama.toLowerCase();
  if (lowerBase.includes(lowerAgama)) {
    return baseMapelName;
  }

  // Indonesian standard Kurikulum Merdeka subject title for religion:
  return `Pendidikan Agama ${cleanAgama} dan Budi Pekerti`;
}

/**
 * Automatically synchronizes a single student's data across:
 * - Data Kelas (jumlahSiswa count & auto-creation of class)
 * - Data Mapel & Pelajaran Agama Sesuai Agama Siswa
 * - Nama Siswa di Nilai (AcademicScore)
 * - Kehadiran (Attendance)
 * - Ekstrakurikuler (Extracurricular)
 * - Catatan Wali Kelas / Kurikuler (TeacherNote)
 * - Keputusan Siswa (StudentDecision)
 * - Piagam & P5
 */
export function syncStudentRelations(
  prevDb: DatabaseState,
  studentInput: Student,
  oldNisn?: string
): DatabaseState {
  const cleanIdSiswa = (studentInput.idSiswa || '').trim() || `S_${Date.now()}`;
  const cleanInputNisn = isPlaceholderNisn(studentInput.nisn)
    ? `BLM_${cleanIdSiswa}`
    : studentInput.nisn.trim();
  const student: Student = {
    ...studentInput,
    idSiswa: cleanIdSiswa,
    nisn: cleanInputNisn
  };

  const targetNisn = student.nisn;
  const previousNisn = oldNisn && !isPlaceholderNisn(oldNisn) ? oldNisn.trim() : targetNisn;
  const semesters: SemesterType[] = ['Ganjil', 'Genap'];

  // 1. DATA SISWA: Pastikan data siswa ada/diupdate di list siswa (tangkap data lama sebelum ditimpa)
  const siswaList = [...prevDb.siswa];
  const sIdx = siswaList.findIndex((s) => s.idSiswa === student.idSiswa);
  const prevStudent = sIdx >= 0 ? siswaList[sIdx] : undefined;
  const resolvedPrevNisn =
    oldNisn && !isPlaceholderNisn(oldNisn)
      ? oldNisn.trim()
      : prevStudent?.nisn && !isPlaceholderNisn(prevStudent.nisn)
      ? prevStudent.nisn
      : targetNisn;
  const effectivePrevNisn = resolvedPrevNisn || targetNisn;

  if (sIdx >= 0) {
    siswaList[sIdx] = student;
  } else {
    siswaList.push(student);
  }

  // 2. DATA KELAS: Hitung otomatis jumlah siswa aktif per kelas & pastikan kelas siswa terdaftar
  const kelasList = [...prevDb.kelas];
  const targetKelasIdx = kelasList.findIndex((k) => k.namaKelas === student.kelas);
  if (targetKelasIdx < 0) {
    // Buat kelas otomatis jika belum ada
    kelasList.push({
      idKelas: `K_${student.kelas}`,
      namaKelas: student.kelas,
      fase: student.fase || 'B',
      waliKelas: prevDb.guru[0]?.namaGuru || 'Wali Kelas',
      tahunPelajaran: prevDb.sekolah.tahunPelajaran,
      semester: prevDb.sekolah.semester,
      jumlahSiswa: 1
    });
  }

  // Perbarui jumlah siswa untuk setiap kelas secara otomatis
  const updatedKelasList = kelasList.map((k) => ({
    ...k,
    jumlahSiswa: siswaList.filter((s) => s.kelas === k.namaKelas && s.statusSiswa === 'Aktif').length
  }));

  // 3. DATA NILAI AKADEMIK & MATA PELAJARAN (SESUAI AGAMA SISWA)
  const nilaiList = [...prevDb.nilaiAkademik];
  semesters.forEach((sem) => {
    prevDb.mataPelajaran.forEach((m) => {
      // Cari nilai berdasarkan idSiswa atau NISN valid (jangan pernah cocokkan NISN placeholder '-')
      const existingIdx = nilaiList.findIndex(
        (n) =>
          ((n.idSiswa && n.idSiswa === student.idSiswa) ||
            (!isPlaceholderNisn(n.nisn) &&
              (n.nisn === effectivePrevNisn || n.nisn === previousNisn || n.nisn === targetNisn))) &&
          n.mataPelajaran === m.mataPelajaran &&
          n.semester === sem
      );

      if (existingIdx >= 0) {
        // Update informasi siswa jika berubah; jangan timpa atau buat nilai otomatis
        const existing = nilaiList[existingIdx];
        const callName = formatStudentCallName(student.namaSiswa, student.bukuInduk?.namaPanggilan);
        if (hasKurikulumMerdekaInputs(existing)) {
          const matchedTps = getMatchedTpsForSubjectAndStudent(
            m.mataPelajaran,
            student.kelas,
            prevDb.tujuanPembelajaran,
            student.agama
          );
          const lmTitles = prevDb.lingkupMateriTitles?.[`${student.kelas}::${m.mataPelajaran}`];
          const km = calculateKurikulumMerdekaScores(
            existing,
            matchedTps,
            m.mataPelajaran,
            lmTitles,
            prevDb.kkmSettings,
            callName
          );
          const finalNa = km.nilaiAkhir > 0 ? km.nilaiAkhir : existing.nilaiAkhir;
          nilaiList[existingIdx] = {
            ...existing,
            idSiswa: student.idSiswa,
            nisn: targetNisn,
            namaSiswa: student.namaSiswa,
            kelas: student.kelas,
            kkm: m.kkm ?? existing.kkm ?? 75,
            naSumatifTp: km.naSumatifTp,
            naSumatifLm: km.naSumatifLm,
            naSumatifSas: km.naSumatifSas,
            jumlahNilai: km.jumlahNilai,
            nilaiAkhir: finalNa,
            predikat: getPredikatFromScore(finalNa, prevDb.kkmSettings?.kktpRanges),
            capaianKompetensi: getCapaianFromScore(finalNa, prevDb.kkmSettings?.kktpRanges),
            deskripsiRapor: km.deskripsiRapor || existing.deskripsiRapor,
            deskripsiTercapai: km.deskripsiTercapai || existing.deskripsiTercapai,
            deskripsiPerluPeningkatan:
              km.deskripsiPerluPeningkatan || existing.deskripsiPerluPeningkatan
          };
          return;
        }

        const summary = calculateScoreSummaryFromFreeInputs(existing, SCORE_INPUT_FIELDS);
        const effectiveNA = summary.filledCount > 0 ? summary.nilaiAkhir : existing.nilaiAkhir;
        const hasScore = summary.filledCount > 0 || effectiveNA > 0;
        const isAgama = m.mataPelajaran.toLowerCase().includes('agama');
        const tpDesc =
          isAgama && hasScore
            ? generateReportDescriptionFromTP(
                m.mataPelajaran,
                student.kelas,
                effectiveNA,
                prevDb.tujuanPembelajaran,
                student.namaSiswa,
                student.agama,
                prevDb.kkmSettings,
                student.bukuInduk?.namaPanggilan
              )
            : null;

        nilaiList[existingIdx] = {
          ...existing,
          idSiswa: student.idSiswa,
          nisn: targetNisn,
          namaSiswa: student.namaSiswa,
          kelas: student.kelas,
          kkm: m.kkm ?? existing.kkm ?? 75,
          jumlahNilai: summary.filledCount > 0 ? summary.jumlahNilai : (existing.jumlahNilai ?? 0),
          deskripsiRapor: hasScore ? (tpDesc ? tpDesc.deskripsiRapor : existing.deskripsiRapor) : '',
          deskripsiTercapai: hasScore
            ? tpDesc
              ? tpDesc.deskripsiTercapai || existing.deskripsiTercapai
              : existing.deskripsiTercapai
            : '',
          deskripsiPerluPeningkatan: hasScore
            ? tpDesc
              ? tpDesc.deskripsiPerluPeningkatan || existing.deskripsiPerluPeningkatan
              : existing.deskripsiPerluPeningkatan
            : ''
        };
      }
    });
  });

  const isOwnedByStudent = (rec: { idSiswa?: string; nisn?: string; namaSiswa?: string; kelas?: string }): boolean => {
    if (rec.idSiswa && rec.idSiswa === student.idSiswa) return true;
    if (
      !isPlaceholderNisn(rec.nisn) &&
      (rec.nisn === effectivePrevNisn || rec.nisn === previousNisn || rec.nisn === targetNisn)
    ) {
      return true;
    }
    const cleanRecName = (rec.namaSiswa || '').trim().toLowerCase();
    if (!cleanRecName) return false;
    if (
      prevStudent &&
      cleanRecName === prevStudent.namaSiswa.trim().toLowerCase() &&
      rec.kelas === prevStudent.kelas &&
      (!rec.idSiswa || rec.idSiswa === student.idSiswa)
    ) {
      return true;
    }
    if (
      cleanRecName === student.namaSiswa.trim().toLowerCase() &&
      rec.kelas === student.kelas &&
      !rec.idSiswa &&
      isPlaceholderNisn(rec.nisn)
    ) {
      return true;
    }
    return false;
  };

  // 4. KEHADIRAN (ATTENDANCE)
  const kehadiranList = [...prevDb.kehadiran];
  semesters.forEach((sem) => {
    const attIdx = kehadiranList.findIndex((h) => isOwnedByStudent(h) && h.semester === sem);
    if (attIdx >= 0) {
      kehadiranList[attIdx] = {
        ...kehadiranList[attIdx],
        idSiswa: student.idSiswa,
        nisn: targetNisn,
        namaSiswa: student.namaSiswa,
        kelas: student.kelas
      };
    } else {
      kehadiranList.push({
        id: `ATT_${targetNisn}_${sem}`,
        idSiswa: student.idSiswa,
        tahunPelajaran: prevDb.sekolah.tahunPelajaran,
        semester: sem,
        kelas: student.kelas,
        nisn: targetNisn,
        namaSiswa: student.namaSiswa,
        sakit: 0,
        izin: 0,
        alpa: 0,
        totalHari: 0
      });
    }
  });

  // 5. CATATAN WALI KELAS (TEACHER NOTE)
  const catatanList = [...prevDb.catatanWali];
  semesters.forEach((sem) => {
    const noteIdx = catatanList.findIndex((c) => isOwnedByStudent(c) && c.semester === sem);
    if (noteIdx >= 0) {
      catatanList[noteIdx] = {
        ...catatanList[noteIdx],
        idSiswa: student.idSiswa,
        nisn: targetNisn,
        namaSiswa: student.namaSiswa,
        kelas: student.kelas
      };
    } else {
      catatanList.push({
        id: `NOTE_${targetNisn}_${sem}`,
        idSiswa: student.idSiswa,
        tahunPelajaran: prevDb.sekolah.tahunPelajaran,
        semester: sem,
        kelas: student.kelas,
        nisn: targetNisn,
        namaSiswa: student.namaSiswa,
        catatanWaliKelas:
          'Menunjukkan sikap belajar yang baik dan aktif dalam kelas. Terus pertahankan semangat belajarmu!'
      });
    }
  });

  // 6. EKSTRAKURIKULER (EXTRACURRICULAR) - pastikan setiap semester punya data
  const ekskulList = [...prevDb.ekstrakurikuler];
  for (let i = 0; i < ekskulList.length; i++) {
    const e = ekskulList[i];
    if (isOwnedByStudent(e)) {
      ekskulList[i] = {
        ...e,
        idSiswa: student.idSiswa,
        nisn: targetNisn,
        namaSiswa: student.namaSiswa,
        kelas: student.kelas
      };
    }
  }
  semesters.forEach((sem) => {
    const hasSemEkskul = ekskulList.some(
      (e) =>
        (e.idSiswa === student.idSiswa || e.nisn === targetNisn || e.nisn === effectivePrevNisn) &&
        e.semester === sem
    );
    if (!hasSemEkskul) {
      ekskulList.push({
        id: `EKS_${targetNisn}_pramuka_${sem}`,
        idSiswa: student.idSiswa,
        tahunPelajaran: prevDb.sekolah.tahunPelajaran,
        semester: sem,
        kelas: student.kelas,
        nisn: targetNisn,
        namaSiswa: student.namaSiswa,
        jenisEkstrakurikuler: 'Pramuka (Wajib)',
        predikat: 'Baik',
        keterangan: 'Aktif dan disiplin dalam mengikuti seluruh kegiatan kepramukaan.'
      });
    }
  });

  // 7. KEPUTUSAN SISWA (STUDENT DECISION)
  const keputusanList = [...(prevDb.keputusanSiswa || [])];
  semesters.forEach((sem) => {
    const decIdx = keputusanList.findIndex((d) => isOwnedByStudent(d) && d.semester === sem);
    if (decIdx >= 0) {
      keputusanList[decIdx] = {
        ...keputusanList[decIdx],
        idSiswa: student.idSiswa,
        nisn: targetNisn,
        namaSiswa: student.namaSiswa,
        kelas: student.kelas
      };
    } else {
      keputusanList.push({
        id: `DEC_${student.idSiswa}_${student.kelas}_${sem}`,
        idSiswa: student.idSiswa,
        nisn: targetNisn,
        namaSiswa: student.namaSiswa,
        kelas: student.kelas,
        semester: sem,
        tahunPelajaran: prevDb.sekolah.tahunPelajaran,
        statusKeputusan: 'Naik Kelas',
        naikKelas: getNextClass(student.kelas),
        tinggalKelas: '-'
      });
    }
  });

  // 8. PIAGAM & P5 (PROJEK, TAHAPAN, DIMENSI, NILAI CAPAIAN P5)
  const piagamList = (prevDb.piagam || []).map((p) =>
    isOwnedByStudent(p)
      ? { ...p, idSiswa: student.idSiswa, nisn: targetNisn, namaSiswa: student.namaSiswa, kelas: student.kelas }
      : p
  );

  // Pastikan kelas siswa memiliki minimal 1 Projek P5 & Tahapan Projek agar nama siswa otomatis masuk P5
  const p5ProjekList = [...(prevDb.p5Projek || [])];
  const p5TahapanList = [...(prevDb.p5Tahapan || [])];
  let projekKelas = p5ProjekList.filter((p) => p.kelas === student.kelas);
  if (projekKelas.length === 0) {
    const autoProjId = `P_${student.kelas}_01`;
    const newProj = {
      idProjek: autoProjId,
      namaProjek: `Projek Penguatan Profil Pelajar Pancasila Kelas ${student.kelas}`,
      tema: 'Kearifan Lokal',
      fase: student.fase || 'B',
      kelas: student.kelas,
      tahunPelajaran: prevDb.sekolah.tahunPelajaran,
      semester: student.semester || prevDb.sekolah.semester || 'Ganjil',
      deskripsiProjek: `Kegiatan projek penguatan karakter Profil Pelajar Pancasila bagi peserta didik Kelas ${student.kelas}.`
    };
    p5ProjekList.push(newProj);
    projekKelas = [newProj];

    const hasTahap = p5TahapanList.some((t) => t.idProjek === autoProjId);
    if (!hasTahap) {
      p5TahapanList.push(
        {
          id: `TPH_${autoProjId}_1`,
          idProjek: autoProjId,
          tahap: 'Pengenalan',
          kegiatan: `Pengenalan tema dan tujuan projek bersama peserta didik Kelas ${student.kelas}.`,
          tanggal: '2026-08-05'
        },
        {
          id: `TPH_${autoProjId}_2`,
          idProjek: autoProjId,
          tahap: 'Aksi',
          kegiatan: `Pelaksanaan aksi nyata projek kolaboratif peserta didik Kelas ${student.kelas}.`,
          tanggal: '2026-10-12'
        }
      );
    }
  }

  const rawP5NilaiList = (prevDb.p5Nilai || []).map((p) => {
    const updated = isOwnedByStudent(p)
      ? { ...p, idSiswa: student.idSiswa, nisn: targetNisn, namaSiswa: student.namaSiswa, kelas: student.kelas }
      : p;
    return {
      ...updated,
      id: buildUniqueP5ScoreId(updated.idProjek, updated.nisn, updated.dimensi)
    };
  });

  // Deduplikasi p5Nilai berdasarkan kombinasi unik (idProjek, idSiswa/nisn, dimensi)
  const p5UniqueMap = new Map<string, (typeof rawP5NilaiList)[number]>();
  rawP5NilaiList.forEach((p) => {
    const ownerKey = p.idSiswa || p.nisn;
    const compositeKey = `${p.idProjek}::${ownerKey}::${(p.dimensi || '').trim().toLowerCase()}`;
    p5UniqueMap.set(compositeKey, p);
  });
  const p5NilaiList = Array.from(p5UniqueMap.values());

  // Gunakan dimensi dari db.p5Dimensi (unik berdasarkan dimensi) atau 3 dimensi utama
  const uniqueDimMap = new Map<string, { dimensi: string; elemen: string }>();
  (prevDb.p5Dimensi || []).forEach((d) => {
    if (d.dimensi && !uniqueDimMap.has(d.dimensi)) {
      uniqueDimMap.set(d.dimensi, { dimensi: d.dimensi, elemen: d.elemen || 'Kolaborasi' });
    }
  });
  const activeDims =
    uniqueDimMap.size > 0
      ? Array.from(uniqueDimMap.values()).slice(0, 6)
      : [
          { dimensi: 'Gotong Royong', elemen: 'Kolaborasi' },
          { dimensi: 'Mandiri', elemen: 'Regulasi Diri' },
          { dimensi: 'Kreatif', elemen: 'Menghasilkan Karya' }
        ];

  projekKelas.forEach((proj) => {
    activeDims.forEach((td) => {
      const existsP5 = p5NilaiList.some(
        (n) =>
          n.idProjek === proj.idProjek &&
          (n.idSiswa === student.idSiswa || (!isPlaceholderNisn(n.nisn) && n.nisn === targetNisn)) &&
          n.dimensi === td.dimensi
      );
      if (!existsP5) {
        p5NilaiList.push({
          id: buildUniqueP5ScoreId(proj.idProjek, targetNisn, td.dimensi),
          idProjek: proj.idProjek,
          idSiswa: student.idSiswa,
          nisn: targetNisn,
          namaSiswa: student.namaSiswa,
          kelas: student.kelas,
          semester: proj.semester || student.semester || 'Ganjil',
          tahunPelajaran: prevDb.sekolah.tahunPelajaran,
          dimensi: td.dimensi,
          elemen: td.elemen,
          capaian: 'BSH',
          deskripsi: `Mampu menunjukkan perkembangan sesuai harapan dalam dimensi ${td.dimensi} pada kegiatan projek.`
        });
      }
    });
  });

  return {
    ...prevDb,
    siswa: siswaList,
    kelas: updatedKelasList,
    nilaiAkademik: nilaiList,
    kehadiran: kehadiranList,
    catatanWali: catatanList,
    ekstrakurikuler: ekskulList,
    keputusanSiswa: keputusanList,
    piagam: piagamList,
    p5Projek: p5ProjekList,
    p5Tahapan: p5TahapanList,
    p5Nilai: p5NilaiList
  };
}

/**
 * Synchronize batch student imports or additions
 */
export function syncBatchStudentsRelations(prevDb: DatabaseState, students: Student[]): DatabaseState {
  let currentDb = ensureUniqueStudentNisns(prevDb);
  for (const s of students) {
    const existing = currentDb.siswa.find((ex) => ex.idSiswa === s.idSiswa);
    currentDb = syncStudentRelations(currentDb, s, existing?.nisn);
  }
  return ensureUniqueStudentNisns(currentDb);
}

/**
 * Cleans up relations when a student is deleted and recounts class students
 */
export function cleanupDeletedStudentRelations(prevDb: DatabaseState, idSiswa: string): DatabaseState {
  const target = prevDb.siswa.find((s) => s.idSiswa === idSiswa);
  if (!target) return prevDb;

  const targetNisn = target.nisn;
  const remainingSiswa = prevDb.siswa.filter((s) => s.idSiswa !== idSiswa);

  const updatedKelas = prevDb.kelas.map((k) => ({
    ...k,
    jumlahSiswa: remainingSiswa.filter((s) => s.kelas === k.namaKelas && s.statusSiswa === 'Aktif').length
  }));

  const keepOtherStudent = <T extends { idSiswa?: string; nisn?: string }>(item: T): boolean => {
    if (item.idSiswa && item.idSiswa === idSiswa) return false;
    if (!isPlaceholderNisn(targetNisn) && item.nisn === targetNisn) return false;
    return true;
  };

  return {
    ...prevDb,
    siswa: remainingSiswa,
    kelas: updatedKelas,
    nilaiAkademik: prevDb.nilaiAkademik.filter(keepOtherStudent),
    kehadiran: prevDb.kehadiran.filter(keepOtherStudent),
    catatanWali: prevDb.catatanWali.filter(keepOtherStudent),
    ekstrakurikuler: prevDb.ekstrakurikuler.filter(keepOtherStudent),
    keputusanSiswa: (prevDb.keputusanSiswa || []).filter(keepOtherStudent),
    piagam: (prevDb.piagam || []).filter(keepOtherStudent),
    p5Nilai: (prevDb.p5Nilai || []).filter(keepOtherStudent)
  };
}

/**
 * Synchronize all existing students across all classes and subjects
 */
export function syncAllStudents(prevDb: DatabaseState): DatabaseState {
  const normalized = ensureUniqueStudentNisns(prevDb);
  return syncBatchStudentsRelations(normalized, normalized.siswa);
}
