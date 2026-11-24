import React from 'react';

export type CenderawasihPaletteMode = 'cenderawasih' | 'gold' | 'navy' | 'emerald';

interface CenderawasihBorderProps {
  theme: CenderawasihPaletteMode;
}

interface ThemePalette {
  outerBorder: string;
  innerBorder: string;
  goldLight: string;
  goldMid: string;
  goldDark: string;
  wingPrimary: string;
  wingSecondary: string;
  headEmerald: string;
  plumeBright: string;
  plumeDeep: string;
  motifBg: string;
  motifStroke: string;
}

const PALETTES: Record<CenderawasihPaletteMode, ThemePalette> = {
  cenderawasih: {
    outerBorder: '#7c2d12',
    innerBorder: '#b45309',
    goldLight: '#fef08a',
    goldMid: '#eab308',
    goldDark: '#a16207',
    wingPrimary: '#991b1b',
    wingSecondary: '#c2410c',
    headEmerald: '#065f46',
    plumeBright: '#facc15',
    plumeDeep: '#d97706',
    motifBg: '#fffbeb',
    motifStroke: '#b45309',
  },
  gold: {
    outerBorder: '#92400e',
    innerBorder: '#d97706',
    goldLight: '#fef9c3',
    goldMid: '#f59e0b',
    goldDark: '#92400e',
    wingPrimary: '#b45309',
    wingSecondary: '#d97706',
    headEmerald: '#047857',
    plumeBright: '#fde047',
    plumeDeep: '#ca8a04',
    motifBg: '#fefce8',
    motifStroke: '#ca8a04',
  },
  navy: {
    outerBorder: '#1e3a8a',
    innerBorder: '#1d4ed8',
    goldLight: '#fef08a',
    goldMid: '#eab308',
    goldDark: '#1e3a8a',
    wingPrimary: '#1e40af',
    wingSecondary: '#3b82f6',
    headEmerald: '#0f766e',
    plumeBright: '#facc15',
    plumeDeep: '#1d4ed8',
    motifBg: '#eff6ff',
    motifStroke: '#1e40af',
  },
  emerald: {
    outerBorder: '#064e3b',
    innerBorder: '#047857',
    goldLight: '#fef08a',
    goldMid: '#eab308',
    goldDark: '#065f46',
    wingPrimary: '#065f46',
    wingSecondary: '#059669',
    headEmerald: '#047857',
    plumeBright: '#facc15',
    plumeDeep: '#059669',
    motifBg: '#ecfdf5',
    motifStroke: '#047857',
  },
};

/**
 * Burung Cenderawasih (Papuan Bird of Paradise) sudut atas.
 * Digambar dalam coordinate space 220x220 menghadap ke kanan/dalam bingkai,
 * dengan ekor bulu kuning emas menjuntai di rel atas & sulur ukiran di rel kiri.
 */
const TopCornerBird: React.FC<{ p: ThemePalette; idPrefix: string; flipX?: boolean }> = ({
  p,
  idPrefix,
  flipX = false,
}) => {
  const gradGold = `${idPrefix}-gold`;
  const gradWing = `${idPrefix}-wing`;
  const gradPlume = `${idPrefix}-plume`;

  return (
    <svg
      viewBox="0 0 220 220"
      className="w-full h-full overflow-visible"
      aria-hidden="true"
    >
      <defs>
        <linearGradient id={gradGold} x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor={p.goldLight} />
          <stop offset="50%" stopColor={p.goldMid} />
          <stop offset="100%" stopColor={p.goldDark} />
        </linearGradient>
        <linearGradient id={gradWing} x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor={p.wingSecondary} />
          <stop offset="100%" stopColor={p.wingPrimary} />
        </linearGradient>
        <linearGradient id={gradPlume} x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" stopColor={p.plumeDeep} />
          <stop offset="55%" stopColor={p.plumeBright} />
          <stop offset="100%" stopColor={p.goldLight} />
        </linearGradient>
      </defs>

      <g transform={flipX ? 'translate(220, 0) scale(-1, 1)' : undefined}>
        {/* Corner Rosette Medallion Backing */}
        <path
          d="M 8 8 L 70 8 C 66 26 42 28 28 42 C 28 56 26 66 8 70 Z"
          fill={p.motifBg}
          stroke={p.outerBorder}
          strokeWidth="2"
        />
        <path
          d="M 12 12 L 54 12 C 44 26 26 44 12 54 Z"
          fill={`url(#${gradGold})`}
          opacity="0.28"
        />

        {/* Sulur Ukiran Papua Vertikal (mengalir ke bawah di sisi kiri) */}
        <path
          d="M 20 68 C 12 96 28 118 18 148 C 12 166 22 184 16 206"
          fill="none"
          stroke={p.goldDark}
          strokeWidth="2.2"
          strokeLinecap="round"
        />
        <path
          d="M 24 72 C 34 88 22 110 30 132 C 34 146 24 166 28 188"
          fill="none"
          stroke={`url(#${gradPlume})`}
          strokeWidth="1.6"
          strokeLinecap="round"
        />
        {/* Spiral daun pakis Papua kecil */}
        <path
          d="M 20 96 C 32 94 35 104 28 108 C 24 110 21 106 24 103"
          fill="none"
          stroke={p.wingPrimary}
          strokeWidth="1.8"
          strokeLinecap="round"
        />
        <path
          d="M 19 138 C 31 136 34 146 27 150 C 23 152 20 148 23 145"
          fill="none"
          stroke={p.wingPrimary}
          strokeWidth="1.6"
          strokeLinecap="round"
        />

        {/* Ekor Filamen Emas Cenderawasih (menjuntai panjang di sisi atas) */}
        <path
          d="M 60 32 C 92 14 134 26 172 16 C 188 12 202 15 212 20"
          fill="none"
          stroke={`url(#${gradPlume})`}
          strokeWidth="3.2"
          strokeLinecap="round"
        />
        <path
          d="M 58 37 C 96 24 132 34 168 23 C 184 18 198 22 206 27"
          fill="none"
          stroke={p.goldMid}
          strokeWidth="2.1"
          strokeLinecap="round"
        />
        <path
          d="M 55 41 C 88 34 124 41 156 30 C 170 26 184 29 194 34"
          fill="none"
          stroke={p.goldLight}
          strokeWidth="1.6"
          strokeLinecap="round"
        />
        {/* Dua helai kawat ekor khas Cenderawasih (twin tail wires) */}
        <path
          d="M 56 44 C 102 48 134 12 178 11 C 194 11 206 17 215 12"
          fill="none"
          stroke={p.outerBorder}
          strokeWidth="1.1"
          strokeLinecap="round"
        />
        <path
          d="M 54 46 C 94 54 128 24 168 38 C 184 43 198 37 210 30"
          fill="none"
          stroke={p.goldDark}
          strokeWidth="1"
          strokeLinecap="round"
        />

        {/* Rumbai bulu sayap emas di belakang badan */}
        <path
          d="M 52 34 C 76 18 106 17 130 22 C 110 27 88 31 64 41 Z"
          fill={`url(#${gradPlume})`}
          opacity="0.92"
        />
        <path
          d="M 50 38 C 72 26 98 25 118 29 C 98 33 78 36 58 43 Z"
          fill={p.goldLight}
          opacity="0.85"
        />

        {/* Ranting ukiran tempat hinggap */}
        <path
          d="M 14 66 Q 36 54 58 58"
          fill="none"
          stroke={p.outerBorder}
          strokeWidth="2.4"
          strokeLinecap="round"
        />
        <circle cx="14" cy="66" r="2.6" fill={p.goldMid} />

        {/* Sayap & Badan Burung Cenderawasih */}
        <path
          d="M 33 52 C 30 40 38 30 50 30 C 60 30 68 37 64 46 C 58 54 42 58 33 52 Z"
          fill={`url(#${gradWing})`}
          stroke={p.outerBorder}
          strokeWidth="1"
        />
        {/* Aksen bulu sayap emas */}
        <path
          d="M 41 42 C 48 36 57 36 62 40 C 54 43 47 46 40 48 Z"
          fill={`url(#${gradGold})`}
        />
        {/* Dada & Leher Zamrud Berkilau */}
        <path
          d="M 29 45 C 27 36 33 27 42 26 C 46 29 45 37 40 44 C 36 49 31 50 29 45 Z"
          fill={p.headEmerald}
        />
        {/* Mahkota Kepala Kuning Emas */}
        <path
          d="M 28 32 C 26 24 34 19 42 22 C 44 23 44 26 42 28 C 36 28 32 30 28 32 Z"
          fill={p.plumeBright}
          stroke={p.goldDark}
          strokeWidth="0.7"
        />
        {/* Paruh Melengkung Khas Cenderawasih */}
        <path
          d="M 28 27 C 21 27 17 30 16 32 C 20 31 24 31 28 30 Z"
          fill="#e2e8f0"
          stroke="#475569"
          strokeWidth="0.7"
        />
        {/* Mata tajam */}
        <circle cx="33" cy="27" r="1.7" fill="#0f172a" />
        <circle cx="32.5" cy="26.5" r="0.6" fill="#ffffff" />

        {/* Cakar kecil bertengger */}
        <path
          d="M 42 54 L 39 59 M 45 54 L 44 59"
          stroke={p.outerBorder}
          strokeWidth="1.2"
          strokeLinecap="round"
        />
      </g>
    </svg>
  );
};

/**
 * Burung Cenderawasih (Papuan Bird of Paradise) sudut bawah — tetap berdiri tegak
 * menghadap ke dalam bingkai dengan bulu ekor menjuntai anggun di rel bawah.
 */
const BottomCornerBird: React.FC<{ p: ThemePalette; idPrefix: string; flipX?: boolean }> = ({
  p,
  idPrefix,
  flipX = false,
}) => {
  const gradGold = `${idPrefix}-gold`;
  const gradWing = `${idPrefix}-wing`;
  const gradPlume = `${idPrefix}-plume`;

  return (
    <svg
      viewBox="0 0 220 220"
      className="w-full h-full overflow-visible"
      aria-hidden="true"
    >
      <defs>
        <linearGradient id={gradGold} x1="0%" y1="100%" x2="100%" y2="0%">
          <stop offset="0%" stopColor={p.goldLight} />
          <stop offset="50%" stopColor={p.goldMid} />
          <stop offset="100%" stopColor={p.goldDark} />
        </linearGradient>
        <linearGradient id={gradWing} x1="0%" y1="0%" x2="100%" y2="100%">
          <stop offset="0%" stopColor={p.wingSecondary} />
          <stop offset="100%" stopColor={p.wingPrimary} />
        </linearGradient>
        <linearGradient id={gradPlume} x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" stopColor={p.plumeDeep} />
          <stop offset="55%" stopColor={p.plumeBright} />
          <stop offset="100%" stopColor={p.goldLight} />
        </linearGradient>
      </defs>

      <g transform={flipX ? 'translate(220, 0) scale(-1, 1)' : undefined}>
        {/* Corner Rosette Medallion Backing */}
        <path
          d="M 8 212 L 70 212 C 66 194 42 192 28 178 C 28 164 26 154 8 150 Z"
          fill={p.motifBg}
          stroke={p.outerBorder}
          strokeWidth="2"
        />
        <path
          d="M 12 208 L 54 208 C 44 194 26 176 12 166 Z"
          fill={`url(#${gradGold})`}
          opacity="0.28"
        />

        {/* Sulur Ukiran Papua Naik di Sisi Kiri */}
        <path
          d="M 20 150 C 12 122 28 100 18 70 C 12 52 22 34 16 14"
          fill="none"
          stroke={p.goldDark}
          strokeWidth="2.2"
          strokeLinecap="round"
        />
        <path
          d="M 24 146 C 34 130 22 108 30 86 C 34 72 24 52 28 30"
          fill="none"
          stroke={`url(#${gradPlume})`}
          strokeWidth="1.6"
          strokeLinecap="round"
        />
        <path
          d="M 20 124 C 32 126 35 116 28 112 C 24 110 21 114 24 117"
          fill="none"
          stroke={p.wingPrimary}
          strokeWidth="1.8"
          strokeLinecap="round"
        />

        {/* Ekor Bulu Emas Cenderawasih Menyapu Rel Bawah */}
        <path
          d="M 60 186 C 94 204 134 192 172 202 C 188 206 202 203 212 198"
          fill="none"
          stroke={`url(#${gradPlume})`}
          strokeWidth="3.2"
          strokeLinecap="round"
        />
        <path
          d="M 58 181 C 96 194 132 184 168 195 C 184 200 198 196 206 191"
          fill="none"
          stroke={p.goldMid}
          strokeWidth="2.1"
          strokeLinecap="round"
        />
        <path
          d="M 55 177 C 88 184 124 177 156 188 C 170 192 184 189 194 184"
          fill="none"
          stroke={p.goldLight}
          strokeWidth="1.6"
          strokeLinecap="round"
        />
        <path
          d="M 56 174 C 102 170 134 206 178 207 C 194 207 206 201 215 206"
          fill="none"
          stroke={p.outerBorder}
          strokeWidth="1.1"
          strokeLinecap="round"
        />

        {/* Rumbai bulu sayap emas bawah */}
        <path
          d="M 52 184 C 76 200 106 201 130 196 C 110 191 88 187 64 177 Z"
          fill={`url(#${gradPlume})`}
          opacity="0.92"
        />

        {/* Ranting ukiran tempat bertengger tegak */}
        <path
          d="M 14 196 Q 36 192 58 196"
          fill="none"
          stroke={p.outerBorder}
          strokeWidth="2.4"
          strokeLinecap="round"
        />
        <circle cx="14" cy="196" r="2.6" fill={p.goldMid} />

        {/* Burung Cenderawasih berdiri tegak di sudut bawah */}
        <path
          d="M 35 186 C 32 174 40 164 52 164 C 62 164 70 171 66 180 C 60 188 44 192 35 186 Z"
          fill={`url(#${gradWing})`}
          stroke={p.outerBorder}
          strokeWidth="1"
        />
        <path
          d="M 43 176 C 50 170 59 170 64 174 C 56 177 49 180 42 182 Z"
          fill={`url(#${gradGold})`}
        />
        <path
          d="M 31 179 C 29 170 35 161 44 160 C 48 163 47 171 42 178 C 38 183 33 184 31 179 Z"
          fill={p.headEmerald}
        />
        <path
          d="M 30 166 C 28 158 36 153 44 156 C 46 157 46 160 44 162 C 38 162 34 164 30 166 Z"
          fill={p.plumeBright}
          stroke={p.goldDark}
          strokeWidth="0.7"
        />
        <path
          d="M 30 161 C 23 161 19 164 18 166 C 22 165 26 165 30 164 Z"
          fill="#e2e8f0"
          stroke="#475569"
          strokeWidth="0.7"
        />
        <circle cx="35" cy="161" r="1.7" fill="#0f172a" />
        <circle cx="34.5" cy="160.5" r="0.6" fill="#ffffff" />
        <path
          d="M 44 188 L 41 194 M 47 188 L 46 194"
          stroke={p.outerBorder}
          strokeWidth="1.2"
          strokeLinecap="round"
        />
      </g>
    </svg>
  );
};

/**
 * Mahkota Ukiran & Bulu Cenderawasih Tengah Atas / Tengah Bawah
 */
const CenterCrest: React.FC<{ p: ThemePalette; inverted?: boolean }> = ({
  p,
  inverted = false,
}) => (
  <svg
    viewBox="0 0 280 44"
    className="w-full h-full overflow-visible"
    aria-hidden="true"
  >
    <g transform={inverted ? 'translate(0, 44) scale(1, -1)' : undefined}>
      {/* Sayap bulu simetris kiri & kanan */}
      <path
        d="M 118 22 C 92 10 56 12 20 22 C 56 20 88 21 116 26 Z"
        fill={p.plumeBright}
        stroke={p.goldDark}
        strokeWidth="0.8"
      />
      <path
        d="M 162 22 C 188 10 224 12 260 22 C 224 20 192 21 164 26 Z"
        fill={p.plumeBright}
        stroke={p.goldDark}
        strokeWidth="0.8"
      />
      <path
        d="M 122 26 C 98 20 72 22 44 29 C 72 28 98 28 120 30 Z"
        fill={p.wingPrimary}
        opacity="0.88"
      />
      <path
        d="M 158 26 C 182 20 208 22 236 29 C 208 28 182 28 160 30 Z"
        fill={p.wingPrimary}
        opacity="0.88"
      />

      {/* Bintang Papua / Ukiran Tengah */}
      <polygon
        points="140,4 146,16 159,22 146,28 140,40 134,28 121,22 134,16"
        fill={p.goldMid}
        stroke={p.outerBorder}
        strokeWidth="1.5"
      />
      <circle
        cx="140"
        cy="22"
        r="5.5"
        fill={p.headEmerald}
        stroke={p.goldLight}
        strokeWidth="1.4"
      />
      <circle cx="140" cy="22" r="2" fill={p.goldLight} />
    </g>
  </svg>
);

/**
 * Bingkai Lengkap Cenderawasih & Ukiran Papua untuk Piagam A4 Landscape.
 * Dirancang berbasis vektor murni agar tajam pada layar dan cetak (window.print).
 */
export const CenderawasihBorder: React.FC<CenderawasihBorderProps> = ({ theme }) => {
  const p = PALETTES[theme];

  return (
    <div
      className="absolute inset-2 sm:inset-3 pointer-events-none z-10 select-none"
      style={{
        WebkitPrintColorAdjust: 'exact',
        printColorAdjust: 'exact',
      }}
    >
      {/* Frame Lapisan Luar & Rel Ukiran Papua Geometri */}
      <div
        className="w-full h-full relative rounded-xs"
        style={{
          border: `3.5px solid ${p.outerBorder}`,
          boxShadow: `inset 0 0 0 2px ${p.goldMid}`,
        }}
      >
        {/* Pita Ukiran Geometri Papuan Atas */}
        <svg
          className="absolute top-[3px] left-[42px] right-[42px] h-[11px] w-[calc(100%-84px)]"
          aria-hidden="true"
        >
          <defs>
            <pattern id="ukiran-h-top" width="24" height="11" patternUnits="userSpaceOnUse">
              <rect width="24" height="11" fill={p.motifBg} />
              <polygon
                points="0,5.5 6,1 12,5.5 6,10"
                fill={p.goldMid}
                fillOpacity="0.45"
                stroke={p.motifStroke}
                strokeWidth="0.9"
              />
              <polygon
                points="12,5.5 18,1 24,5.5 18,10"
                fill={p.wingPrimary}
                fillOpacity="0.22"
                stroke={p.motifStroke}
                strokeWidth="0.9"
              />
              <circle cx="6" cy="5.5" r="1.1" fill={p.outerBorder} />
              <circle cx="18" cy="5.5" r="1.1" fill={p.headEmerald} />
            </pattern>
          </defs>
          <rect width="100%" height="11" fill="url(#ukiran-h-top)" />
        </svg>

        {/* Pita Ukiran Geometri Papuan Bawah */}
        <svg
          className="absolute bottom-[3px] left-[42px] right-[42px] h-[11px] w-[calc(100%-84px)]"
          aria-hidden="true"
        >
          <defs>
            <pattern id="ukiran-h-bot" width="24" height="11" patternUnits="userSpaceOnUse">
              <rect width="24" height="11" fill={p.motifBg} />
              <polygon
                points="0,5.5 6,1 12,5.5 6,10"
                fill={p.goldMid}
                fillOpacity="0.45"
                stroke={p.motifStroke}
                strokeWidth="0.9"
              />
              <polygon
                points="12,5.5 18,1 24,5.5 18,10"
                fill={p.wingPrimary}
                fillOpacity="0.22"
                stroke={p.motifStroke}
                strokeWidth="0.9"
              />
              <circle cx="6" cy="5.5" r="1.1" fill={p.outerBorder} />
              <circle cx="18" cy="5.5" r="1.1" fill={p.headEmerald} />
            </pattern>
          </defs>
          <rect width="100%" height="11" fill="url(#ukiran-h-bot)" />
        </svg>

        {/* Pita Ukiran Geometri Papuan Kiri */}
        <svg
          className="absolute left-[3px] top-[42px] bottom-[42px] w-[11px] h-[calc(100%-84px)]"
          aria-hidden="true"
        >
          <defs>
            <pattern id="ukiran-v-left" width="11" height="24" patternUnits="userSpaceOnUse">
              <rect width="11" height="24" fill={p.motifBg} />
              <polygon
                points="5.5,0 10,6 5.5,12 1,6"
                fill={p.goldMid}
                fillOpacity="0.45"
                stroke={p.motifStroke}
                strokeWidth="0.9"
              />
              <polygon
                points="5.5,12 10,18 5.5,24 1,18"
                fill={p.wingPrimary}
                fillOpacity="0.22"
                stroke={p.motifStroke}
                strokeWidth="0.9"
              />
              <circle cx="5.5" cy="6" r="1.1" fill={p.outerBorder} />
              <circle cx="5.5" cy="18" r="1.1" fill={p.headEmerald} />
            </pattern>
          </defs>
          <rect width="11" height="100%" fill="url(#ukiran-v-left)" />
        </svg>

        {/* Pita Ukiran Geometri Papuan Kanan */}
        <svg
          className="absolute right-[3px] top-[42px] bottom-[42px] w-[11px] h-[calc(100%-84px)]"
          aria-hidden="true"
        >
          <defs>
            <pattern id="ukiran-v-right" width="11" height="24" patternUnits="userSpaceOnUse">
              <rect width="11" height="24" fill={p.motifBg} />
              <polygon
                points="5.5,0 10,6 5.5,12 1,6"
                fill={p.goldMid}
                fillOpacity="0.45"
                stroke={p.motifStroke}
                strokeWidth="0.9"
              />
              <polygon
                points="5.5,12 10,18 5.5,24 1,18"
                fill={p.wingPrimary}
                fillOpacity="0.22"
                stroke={p.motifStroke}
                strokeWidth="0.9"
              />
              <circle cx="5.5" cy="6" r="1.1" fill={p.outerBorder} />
              <circle cx="5.5" cy="18" r="1.1" fill={p.headEmerald} />
            </pattern>
          </defs>
          <rect width="11" height="100%" fill="url(#ukiran-v-right)" />
        </svg>

        {/* Garis Bingkai Dalam Tipis */}
        <div
          className="absolute inset-[15px] rounded-xs"
          style={{
            border: `1.5px solid ${p.innerBorder}`,
            boxShadow: `inset 0 0 0 1px ${p.goldLight}`,
          }}
        />

        {/* Ornamen 4 Sudut Burung Cenderawasih */}
        <div className="absolute -top-2 -left-2 w-40 h-40 sm:w-44 sm:h-44">
          <TopCornerBird p={p} idPrefix="tl" />
        </div>
        <div className="absolute -top-2 -right-2 w-40 h-40 sm:w-44 sm:h-44">
          <TopCornerBird p={p} idPrefix="tr" flipX />
        </div>
        <div className="absolute -bottom-2 -left-2 w-40 h-40 sm:w-44 sm:h-44">
          <BottomCornerBird p={p} idPrefix="bl" />
        </div>
        <div className="absolute -bottom-2 -right-2 w-40 h-40 sm:w-44 sm:h-44">
          <BottomCornerBird p={p} idPrefix="br" flipX />
        </div>

        {/* Mahkota Ukiran & Bulu Cenderawasih Tengah Atas */}
        <div className="hidden sm:block absolute -top-2.5 left-1/2 -translate-x-1/2 w-56 h-9">
          <CenterCrest p={p} />
        </div>

        {/* Ornamen Ukiran & Bulu Cenderawasih Tengah Bawah */}
        <div className="hidden sm:block absolute -bottom-2.5 left-1/2 -translate-x-1/2 w-48 h-8">
          <CenterCrest p={p} inverted />
        </div>
      </div>
    </div>
  );
};
