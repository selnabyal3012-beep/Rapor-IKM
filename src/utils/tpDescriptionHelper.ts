import { LearningObjective, ClassLevel, KkmGlobalSettings, KktpRangeItem } from '../types';

const FALLBACK_KKTP_RANGES: KktpRangeItem[] = [
  {
    id: 'sangat_baik',
    min: 88,
    max: 100,
    predikat: 'Sangat Baik',
    labelPredikat: 'Sangat baik',
    kalimatAwal: 'Menunjukkan penguasaan yang sangat baik dalam'
  },
  {
    id: 'baik',
    min: 74,
    max: 87,
    predikat: 'Baik',
    labelPredikat: 'Baik',
    kalimatAwal: 'Menunjukkan penguasaan yang baik dalam'
  },
  {
    id: 'cukup',
    min: 60,
    max: 73,
    predikat: 'Cukup',
    labelPredikat: 'Cukup',
    kalimatAwal: 'Menunjukkan penguasaan yang cukup dalam'
  },
  {
    id: 'perlu_bimbingan',
    min: 0,
    max: 59,
    predikat: 'Perlu Bimbingan',
    labelPredikat: 'Perlu bimbingan',
    kalimatAwal: 'Perlu bimbingan dalam'
  }
];

export function formatStudentCallName(namaSiswa?: string, namaPanggilan?: string): string {
  const nick = (namaPanggilan || '').trim();
  if (nick) {
    return nick.charAt(0).toUpperCase() + nick.slice(1).toLowerCase();
  }
  const full = (namaSiswa || '').trim();
  if (!full) return '';
  const firstToken = full.split(/[\s.]+/).find((t) => t.length >= 2) || full.split(/\s+/)[0] || '';
  if (!firstToken) return '';
  return firstToken.charAt(0).toUpperCase() + firstToken.slice(1).toLowerCase();
}

export function resolveKktpRangeForScore(
  nilai: number,
  kktpRanges?: KktpRangeItem[]
): KktpRangeItem {
  const ranges = (kktpRanges && kktpRanges.length > 0 ? kktpRanges : FALLBACK_KKTP_RANGES)
    .slice()
    .sort((a, b) => b.min - a.min);
  const clamped = Math.min(100, Math.max(0, Math.round(nilai)));
  const found = ranges.find((r) => clamped >= r.min && clamped <= r.max);
  if (found) return found;
  return ranges.find((r) => clamped >= r.min) || ranges[ranges.length - 1];
}

export function buildKktpDescriptionSentence(
  scoreVal: number,
  phrase: string,
  kkmSettings?: KkmGlobalSettings,
  callName?: string
): string {
  const cleanPhrase = (phrase || '').trim().replace(/\.+$/, '');
  if (!cleanPhrase) return '';

  const useRangePredikat = kkmSettings?.deskripsiDenganRentangPredikat ?? true;
  const useNickname = kkmSettings?.deskripsiDiawaliNamaPanggilan ?? true;
  const matchedRange = resolveKktpRangeForScore(scoreVal, kkmSettings?.kktpRanges);

  const cukupMin =
    kkmSettings?.kktpRanges?.find((r) => r.id === 'cukup')?.min ?? 60;

  let prefix = useRangePredikat
    ? matchedRange.kalimatAwal.trim()
    : scoreVal < cukupMin
    ? 'Perlu bimbingan dalam'
    : 'Menguasai dalam';

  if (useNickname && callName) {
    const lowerPrefix = prefix.charAt(0).toLowerCase() + prefix.slice(1);
    return `${callName} ${lowerPrefix} ${cleanPhrase}.`;
  }

  return `${prefix} ${cleanPhrase}.`;
}

/**
 * Clean up common teacher prefixes from TP sentences so they integrate
 * smoothly into official report card sentences (e.g. "Siswa mampu memahami..." -> "memahami...")
 */
export function cleanTpPhrase(text: string): string {
  if (!text) return '';
  let cleaned = text.trim();

  // Strip trailing periods
  cleaned = cleaned.replace(/\.+$/, '');

  // Regex patterns to remove Indonesian pedagogical prefixes
  const prefixRegexes = [
    /^(peserta\s+didik\s+mampu\s+)/i,
    /^(peserta\s+didik\s+dapat\s+)/i,
    /^(siswa\s+mampu\s+)/i,
    /^(siswa\s+dapat\s+)/i,
    /^(murid\s+mampu\s+)/i,
    /^(anak\s+mampu\s+)/i,
    /^(mampu\s+)/i,
    /^(dapat\s+)/i,
    /^(diharapkan\s+mampu\s+)/i,
    /^(mampu\s+menunjukkan\s+)/i
  ];

  for (const regex of prefixRegexes) {
    if (regex.test(cleaned)) {
      cleaned = cleaned.replace(regex, '');
      break;
    }
  }

  cleaned = cleaned.trim();
  if (cleaned.length > 0) {
    // lowercase first character if not an acronym or proper noun
    const firstWord = cleaned.split(' ')[0];
    if (firstWord !== firstWord.toUpperCase()) {
      cleaned = cleaned.charAt(0).toLowerCase() + cleaned.slice(1);
    }
  }

  return cleaned;
}

export interface TpDescriptionResult {
  deskripsiRapor: string;
  deskripsiTercapai: string;
  deskripsiPerluPeningkatan: string;
  linkedTpKodes: string[];
  matchedTps: LearningObjective[];
}

export function getMatchedTpsForSubjectAndStudent(
  mapelName: string,
  kelas: ClassLevel | string,
  allTPs: LearningObjective[],
  agamaSiswa?: string
): LearningObjective[] {
  const normAgama = agamaSiswa ? agamaSiswa.trim().toLowerCase() : '';
  const isAgamaSubject = mapelName.toLowerCase().includes('agama');

  let matchedTps: LearningObjective[] = [];

  if (isAgamaSubject && normAgama) {
    matchedTps = allTPs.filter((tp) => {
      const tpAgama = (tp.agama || '').toLowerCase();
      const tpMapel = tp.mataPelajaran.toLowerCase();
      const kelasMatch = !kelas || tp.kelas === kelas;
      if (!kelasMatch) return false;

      return (
        tpAgama === normAgama ||
        tpMapel.includes(normAgama) ||
        (normAgama === 'kristen' && (tpMapel.includes('kristen') || tpAgama === 'kristen' || tpMapel.includes('protestan'))) ||
        (normAgama === 'katolik' && (tpMapel.includes('katolik') || tpAgama === 'katolik')) ||
        (normAgama === 'islam' && (tpMapel.includes('islam') || tpAgama === 'islam')) ||
        (normAgama === 'hindu' && (tpMapel.includes('hindu') || tpAgama === 'hindu')) ||
        (normAgama === 'buddha' && (tpMapel.includes('buddha') || tpAgama === 'buddha')) ||
        ((normAgama === 'khonghucu' || normAgama === 'konghucu') &&
          (tpMapel.includes('khonghucu') || tpMapel.includes('konghucu') || tpAgama === 'khonghucu' || tpAgama === 'konghucu'))
      );
    });
  }

  if (matchedTps.length === 0) {
    matchedTps = allTPs.filter((tp) => {
      const mapelMatch =
        tp.mataPelajaran.trim().toLowerCase() === mapelName.trim().toLowerCase() ||
        tp.mataPelajaran.toLowerCase().includes(mapelName.toLowerCase()) ||
        mapelName.toLowerCase().includes(tp.mataPelajaran.toLowerCase());

      const kelasMatch = !kelas || tp.kelas === kelas;
      return mapelMatch && kelasMatch;
    });
  }

  return matchedTps;
}

/**
 * Generates official Kurikulum Merdeka report card descriptions dynamically
 * linked to the active Tujuan Pembelajaran (TP) for a specific subject and class.
 */
export function generateReportDescriptionFromTP(
  mapelName: string,
  kelas: ClassLevel | string,
  nilaiAkhir: number,
  allTPs: LearningObjective[],
  namaSiswa?: string,
  agamaSiswa?: string,
  kkmSettings?: KkmGlobalSettings,
  studentNickname?: string
): TpDescriptionResult {
  const matchedTps = getMatchedTpsForSubjectAndStudent(mapelName, kelas, allTPs, agamaSiswa);
  const linkedTpKodes = matchedTps.map((tp) => tp.kodeTp);
  const callName = formatStudentCallName(namaSiswa, studentNickname);

  // If no TP exists yet for this class, fallback to clean subject-based phrases
  if (matchedTps.length === 0) {
    const genericTp1 = `memahami materi inti pembelajaran ${mapelName}`;
    const genericTp2 = `mempraktikkan keterampilan dasar pembelajaran ${mapelName}`;
    const tercapai = buildKktpDescriptionSentence(nilaiAkhir, genericTp1, kkmSettings, callName);
    const perlu = buildKktpDescriptionSentence(nilaiAkhir, genericTp2, kkmSettings, callName);
    const perluNoName = buildKktpDescriptionSentence(nilaiAkhir, genericTp2, kkmSettings, undefined);

    return {
      deskripsiRapor: `${tercapai} ${perluNoName}`.trim(),
      deskripsiTercapai: tercapai,
      deskripsiPerluPeningkatan: perlu,
      linkedTpKodes: [],
      matchedTps: []
    };
  }

  const phrase1 = cleanTpPhrase(matchedTps[0].tujuanPembelajaran);
  const phrase2 = matchedTps[1]
    ? cleanTpPhrase(matchedTps[1].tujuanPembelajaran)
    : matchedTps.length > 2
    ? cleanTpPhrase(matchedTps[2].tujuanPembelajaran)
    : '';

  const deskripsiTercapai = buildKktpDescriptionSentence(nilaiAkhir, phrase1, kkmSettings, callName);
  const deskripsiPerluPeningkatan = phrase2
    ? buildKktpDescriptionSentence(nilaiAkhir, phrase2, kkmSettings, callName)
    : '';
  const deskripsiPerluNoName = phrase2
    ? buildKktpDescriptionSentence(nilaiAkhir, phrase2, kkmSettings, undefined)
    : '';
  const deskripsiRapor = deskripsiPerluNoName
    ? `${deskripsiTercapai} ${deskripsiPerluNoName}`
    : deskripsiTercapai;

  return {
    deskripsiRapor,
    deskripsiTercapai,
    deskripsiPerluPeningkatan,
    linkedTpKodes,
    matchedTps
  };
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

export function getPredikatFromScore(
  nilai: number,
  kktpRanges?: KktpRangeItem[]
): 'Sangat Baik' | 'Baik' | 'Cukup' | 'Perlu Bimbingan' {
  return resolveKktpRangeForScore(nilai, kktpRanges).predikat;
}

export function getCapaianFromScore(
  nilai: number,
  kktpRanges?: KktpRangeItem[]
): 'Sangat Baik' | 'Baik' | 'Cukup' | 'Perlu Pendampingan' {
  const predikat = resolveKktpRangeForScore(nilai, kktpRanges).predikat;
  if (predikat === 'Perlu Bimbingan') return 'Perlu Pendampingan';
  return predikat;
}

export function clampScore(val: any): number {
  const num = Number(val);
  if (isNaN(num)) return 0;
  return Math.min(100, Math.max(0, Math.round(num)));
}

export const SCORE_INPUT_FIELDS = [
  'formatif',
  'sumatif',
  'ph1',
  'ph2',
  'ph3',
  'pts',
  'pas',
  'praktik',
] as const;

export type ScoreInputField = (typeof SCORE_INPUT_FIELDS)[number];

export function calculateScoreSummaryFromFreeInputs(
  score: Partial<Record<ScoreInputField, number | undefined>>,
  fields: readonly ScoreInputField[] = SCORE_INPUT_FIELDS
): { jumlahNilai: number; nilaiAkhir: number; filledCount: number } {
  const filled = fields
    .map((key) => score[key])
    .filter((v): v is number => typeof v === 'number' && !Number.isNaN(v));
  const jumlahNilai = filled.reduce((acc, v) => acc + v, 0);
  const nilaiAkhir = filled.length > 0 ? clampScore(Math.round(jumlahNilai / filled.length)) : 0;
  return { jumlahNilai, nilaiAkhir, filledCount: filled.length };
}

export const TP_COLUMN_KEYS = [
  'tp1', 'tp2', 'tp3', 'tp4', 'tp5', 'tp6',
  'tp7', 'tp8', 'tp9', 'tp10', 'tp11', 'tp12'
] as const;

export const LM_COLUMN_KEYS = [
  'lm1', 'lm2', 'lm3', 'lm4', 'lm5', 'lm6', 'lm7', 'lm8'
] as const;

export interface KurikulumMerdekaScoreInput {
  tpScores?: Record<string, number | undefined>;
  lmScores?: Record<string, number | undefined>;
  sasNonTes?: number;
  sasTes?: number;
}

export function hasKurikulumMerdekaInputs(score: KurikulumMerdekaScoreInput): boolean {
  const tpCount = Object.values(score.tpScores || {}).filter(
    (v): v is number => typeof v === 'number' && !Number.isNaN(v)
  ).length;
  const lmCount = Object.values(score.lmScores || {}).filter(
    (v): v is number => typeof v === 'number' && !Number.isNaN(v)
  ).length;
  const hasSas =
    (typeof score.sasNonTes === 'number' && !Number.isNaN(score.sasNonTes)) ||
    (typeof score.sasTes === 'number' && !Number.isNaN(score.sasTes));
  return tpCount > 0 || lmCount > 0 || hasSas;
}


export function extractTpSlotIndex(kodeTp?: string): number | null {
  if (!kodeTp) return null;
  const match = kodeTp.trim().match(/^(?:tp[\s_-]*)?0*(\d{1,2})$/i);
  if (!match) return null;
  const num = parseInt(match[1], 10);
  if (num >= 1 && num <= TP_COLUMN_KEYS.length) {
    return num - 1;
  }
  return null;
}

export function resolveTpSlots(matchedTps: LearningObjective[]): (LearningObjective | undefined)[] {
  const slots: (LearningObjective | undefined)[] = Array(TP_COLUMN_KEYS.length).fill(undefined);
  const leftovers: LearningObjective[] = [];

  matchedTps.forEach((tp) => {
    const slotIdx = extractTpSlotIndex(tp.kodeTp);
    if (slotIdx !== null && !slots[slotIdx]) {
      slots[slotIdx] = tp;
    } else {
      leftovers.push(tp);
    }
  });

  leftovers.forEach((tp) => {
    const emptyIdx = slots.findIndex((s) => s === undefined);
    if (emptyIdx >= 0) {
      slots[emptyIdx] = tp;
    }
  });

  return slots;
}

export function calculateKurikulumMerdekaScores(
  score: KurikulumMerdekaScoreInput & Partial<Record<ScoreInputField, number | undefined>>,
  matchedTps: LearningObjective[],
  mapelName: string,
  lingkupMateriTitles?: Record<string, string>,
  kkmSettings?: KkmGlobalSettings,
  studentCallName?: string
): {
  naSumatifTp: number | undefined;
  naSumatifLm: number | undefined;
  naSumatifSas: number | undefined;
  nilaiAkhir: number;
  jumlahNilai: number;
  filledCount: number;
  hasKmInputs: boolean;
  deskripsiTercapai: string;
  deskripsiPerluPeningkatan: string;
  deskripsiRapor: string;
} {
  const tpEntries: { key: string; idx: number; val: number }[] = [];
  TP_COLUMN_KEYS.forEach((key, idx) => {
    const raw = score.tpScores?.[key];
    if (typeof raw === 'number' && !Number.isNaN(raw)) {
      tpEntries.push({ key, idx, val: raw });
    }
  });

  const lmEntries: { key: string; idx: number; val: number }[] = [];
  const lmValues: number[] = [];
  LM_COLUMN_KEYS.forEach((key, idx) => {
    const raw = score.lmScores?.[key];
    if (typeof raw === 'number' && !Number.isNaN(raw)) {
      lmEntries.push({ key, idx, val: raw });
      lmValues.push(raw);
    }
  });

  const sasValues: number[] = [];
  if (typeof score.sasNonTes === 'number' && !Number.isNaN(score.sasNonTes)) {
    sasValues.push(score.sasNonTes);
  }
  if (typeof score.sasTes === 'number' && !Number.isNaN(score.sasTes)) {
    sasValues.push(score.sasTes);
  }

  const rawAvgTp =
    tpEntries.length > 0
      ? tpEntries.reduce((acc, item) => acc + item.val, 0) / tpEntries.length
      : undefined;
  const rawAvgLm =
    lmValues.length > 0
      ? lmValues.reduce((acc, v) => acc + v, 0) / lmValues.length
      : undefined;
  const rawAvgSas =
    sasValues.length > 0
      ? sasValues.reduce((acc, v) => acc + v, 0) / sasValues.length
      : undefined;

  const naSumatifTp = rawAvgTp !== undefined ? Number(rawAvgTp.toFixed(2)) : undefined;
  const naSumatifLm = rawAvgLm !== undefined ? Number(rawAvgLm.toFixed(2)) : undefined;
  const naSumatifSas = rawAvgSas !== undefined ? Number(rawAvgSas.toFixed(2)) : undefined;

  // Hitung NA RAPOR berdasarkan Bobot Nilai Rapor (S.TP, S.LM, S.AS)
  const bobotTp =
    kkmSettings?.bobot?.namaPenilaianTp === 'FORMATIF'
      ? 0
      : kkmSettings?.bobot?.bobotTp ?? 1;
  const bobotLm = kkmSettings?.bobot?.bobotLm ?? 1;
  const bobotSas = kkmSettings?.bobot?.bobotSas ?? 2;

  const weightedPairs: { avg: number; weight: number }[] = [];
  if (rawAvgTp !== undefined && bobotTp > 0) weightedPairs.push({ avg: rawAvgTp, weight: bobotTp });
  if (rawAvgLm !== undefined && bobotLm > 0) weightedPairs.push({ avg: rawAvgLm, weight: bobotLm });
  if (rawAvgSas !== undefined && bobotSas > 0) weightedPairs.push({ avg: rawAvgSas, weight: bobotSas });

  const totalActiveWeight = weightedPairs.reduce((acc, p) => acc + p.weight, 0);
  const weightedSum = weightedPairs.reduce((acc, p) => acc + p.avg * p.weight, 0);

  const kmFilledCount = tpEntries.length + lmValues.length + sasValues.length;
  const kmTotalSum =
    tpEntries.reduce((acc, i) => acc + i.val, 0) +
    lmValues.reduce((acc, v) => acc + v, 0) +
    sasValues.reduce((acc, v) => acc + v, 0);

  const hasKm = kmFilledCount > 0;
  const legacySummary = calculateScoreSummaryFromFreeInputs(score, SCORE_INPUT_FIELDS);

  const nilaiAkhir = hasKm
    ? totalActiveWeight > 0
      ? clampScore(Math.round(weightedSum / totalActiveWeight))
      : 0
    : legacySummary.nilaiAkhir;

  let deskripsiTercapai = '';
  let deskripsiPerluPeningkatan = '';
  let deskripsiRapor = '';

  const tpSlots = resolveTpSlots(matchedTps);

  if (tpEntries.length > 0) {
    let highest = tpEntries[0];
    let lowest = tpEntries[0];
    tpEntries.forEach((entry) => {
      if (entry.val > highest.val) highest = entry;
      if (entry.val < lowest.val) lowest = entry;
    });

    if (tpEntries.length > 1 && highest.key === lowest.key) {
      const distinctTp = [...tpEntries]
        .reverse()
        .find((e) => e.key !== highest.key && tpSlots[e.idx]?.tujuanPembelajaran);
      lowest = distinctTp || tpEntries[tpEntries.length - 1];
    }

    const resolveTpPhrase = (entry: { idx: number }) => {
      const tpObj = tpSlots[entry.idx] || matchedTps[entry.idx];
      if (tpObj?.tujuanPembelajaran) {
        return cleanTpPhrase(tpObj.tujuanPembelajaran);
      }
      const fallbackLm = lingkupMateriTitles?.[`lm${entry.idx + 1}`];
      if (fallbackLm && fallbackLm.trim()) {
        return `materi ${fallbackLm.trim()} pada ${mapelName}`;
      }
      return `menguasai materi tujuan pembelajaran ${entry.idx + 1} pada ${mapelName}`;
    };

    const highestPhrase = resolveTpPhrase(highest);
    const lowestPhrase = resolveTpPhrase(lowest);

    deskripsiTercapai = buildKktpDescriptionSentence(
      highest.val,
      highestPhrase,
      kkmSettings,
      studentCallName
    );
    deskripsiPerluPeningkatan = buildKktpDescriptionSentence(
      lowest.val,
      lowestPhrase,
      kkmSettings,
      studentCallName
    );

    const hasSecondClause =
      tpEntries.length > 1 && highest.key !== lowest.key && highestPhrase !== lowestPhrase;

    if (hasSecondClause) {
      const lowestWithoutName = buildKktpDescriptionSentence(
        lowest.val,
        lowestPhrase,
        kkmSettings,
        undefined
      );
      deskripsiRapor = `${deskripsiTercapai} ${lowestWithoutName}`;
    } else {
      deskripsiRapor = deskripsiTercapai;
    }
  } else if (lmEntries.length > 0) {
    let highestLm = lmEntries[0];
    let lowestLm = lmEntries[0];
    lmEntries.forEach((entry) => {
      if (entry.val > highestLm.val) highestLm = entry;
      if (entry.val < lowestLm.val) lowestLm = entry;
    });

    const resolveLmPhrase = (entry: { key: string; idx: number }) => {
      const lmName = lingkupMateriTitles?.[entry.key];
      if (lmName && lmName.trim()) {
        return `memahami materi ${lmName.trim()}`;
      }
      const tpObj = tpSlots[entry.idx] || matchedTps[entry.idx];
      if (tpObj?.tujuanPembelajaran) {
        return cleanTpPhrase(tpObj.tujuanPembelajaran);
      }
      return `menguasai lingkup materi ${entry.idx + 1} pada ${mapelName}`;
    };

    const highLmPhrase = resolveLmPhrase(highestLm);
    const lowLmPhrase = resolveLmPhrase(lowestLm);

    deskripsiTercapai = buildKktpDescriptionSentence(
      highestLm.val,
      highLmPhrase,
      kkmSettings,
      studentCallName
    );
    deskripsiPerluPeningkatan = buildKktpDescriptionSentence(
      lowestLm.val,
      lowLmPhrase,
      kkmSettings,
      studentCallName
    );

    if (lmEntries.length > 1 && highestLm.key !== lowestLm.key && highLmPhrase !== lowLmPhrase) {
      const lowWithoutName = buildKktpDescriptionSentence(
        lowestLm.val,
        lowLmPhrase,
        kkmSettings,
        undefined
      );
      deskripsiRapor = `${deskripsiTercapai} ${lowWithoutName}`;
    } else {
      deskripsiRapor = deskripsiTercapai;
    }
  }

  return {
    naSumatifTp,
    naSumatifLm,
    naSumatifSas,
    nilaiAkhir,
    jumlahNilai: hasKm ? kmTotalSum : legacySummary.jumlahNilai,
    filledCount: hasKm ? kmFilledCount : legacySummary.filledCount,
    hasKmInputs: hasKm,
    deskripsiTercapai,
    deskripsiPerluPeningkatan,
    deskripsiRapor
  };
}


