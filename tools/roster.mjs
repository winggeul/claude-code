/**
 * 담당자 → 지점. 집계기와 백필기가 같은 기준을 쓰도록 한곳에 둔다.
 *
 * 통합고객목록의 담당자 열에서 지점을 가른다. 열 위치는 내보낼 때마다 달라져서
 * (CSV 는 5번째, xlsx 는 맨 오른쪽) 자리로 찾지 않고 이 명단과 맞춰 찾는다.
 */

export const SIHEUNG = "시흥";
export const CHEONAN = "천안";

const ROSTER = new Map();
const add = (branch, names) => names.forEach(n => ROSTER.set(n, branch));

add(SIHEUNG, ["최성훈", "백지연", "성예훈", "이민철", "박윤철", "김시원",
              "양지애", "김대식", "전하영", "이지은", "최시형", "김현정"]);
add(CHEONAN, ["홍혜리", "김경태", "조항찬", "서정인", "김수현", "송현정",
              "김유리", "조항준", "박소라", "천안영업팀", "조광연", "노창훈"]);

/** 명단에 있는 이름인가. 담당자 열을 찾을 때 쓴다. */
export const isAgent = (v) => ROSTER.has((v ?? "").trim());

/**
 * 어느 쪽에도 속하지 않는 담당자(미사용, 관리자, 빈칸)를 5 대 11 로 나눈다.
 * 사람 수 비율대로 돌려 담아서 시흥 + 천안이 언제나 합계와 딱 맞게 한다.
 * 배분 대상은 지금까지 하루 한두 건 수준이라 순서가 어떻든 총합에 영향이 없다.
 */
const RATIO = Array.from({ length: 16 }, (_, i) => (i % 3 === 0 && i < 15) ? SIHEUNG : CHEONAN);

export function makeSplitter() {
  let n = 0;
  return (name) => {
    const hit = ROSTER.get((name ?? "").trim());
    if (hit) return hit;
    return RATIO[n++ % RATIO.length];
  };
}

/**
 * 담당자 열의 자리. 헤더가 있으면 이름으로, 없으면 명단과 맞는 비율이 가장 높은 열로 찾는다.
 * rows 는 헤더를 뺀 본문이어야 한다.
 */
export function findAgentColumn(rows, header) {
  if (header) {
    const i = header.findIndex(v => (v ?? "").trim() === "담당자");
    if (i >= 0) return i;
  }
  const width = Math.max(0, ...rows.map(r => r.length));
  let col = -1, best = 0;
  for (let i = 0; i < width; i++) {
    const vals = rows.map(r => (r[i] ?? "").trim()).filter(Boolean);
    if (vals.length < rows.length * 0.5) continue;
    const rate = vals.filter(isAgent).length / vals.length;
    if (rate > best && rate > 0.5) { best = rate; col = i; }
  }
  return col;
}

export const BRANCHES = [SIHEUNG, CHEONAN];
