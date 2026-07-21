import { mkdir, writeFile } from "node:fs/promises";

const username = process.env.GITHUB_USERNAME || "unseenap";
const token = process.env.GITHUB_TOKEN;
const apiHeaders = {
  Accept: "application/vnd.github+json",
  "User-Agent": "unseenap-profile-card-generator",
  ...(token ? { Authorization: `Bearer ${token}` } : {}),
};

async function request(url, headers = apiHeaders) {
  const response = await fetch(url, { headers });
  if (!response.ok) {
    throw new Error(`${response.status} ${response.statusText}: ${url}`);
  }
  return response;
}

async function getRepositories() {
  const repositories = [];
  for (let page = 1; ; page += 1) {
    const batch = await request(
      `https://api.github.com/users/${username}/repos?per_page=100&page=${page}&type=owner`,
    ).then((response) => response.json());
    repositories.push(...batch);
    if (batch.length < 100) return repositories;
  }
}

function isoDate(date) {
  return date.toISOString().slice(0, 10);
}

function xml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

async function getContributionStats() {
  const end = new Date();
  end.setUTCHours(0, 0, 0, 0);
  const start = new Date(end);
  start.setUTCDate(start.getUTCDate() - 364);

  const html = await request(
    `https://github.com/users/${username}/contributions?from=${isoDate(start)}&to=${isoDate(end)}`,
    { "User-Agent": "unseenap-profile-card-generator" },
  ).then((response) => response.text());

  const days = new Map();
  const dayPattern = /data-date="(\d{4}-\d{2}-\d{2})" id="(contribution-day-component-[^"]+)"/g;
  for (const match of html.matchAll(dayPattern)) {
    const date = match[1];
    if (date >= isoDate(start) && date <= isoDate(end)) {
      days.set(match[2], { date, count: 0 });
    }
  }

  const countPattern = /<tool-tip[^>]*for="(contribution-day-component-[^"]+)"[^>]*>(\d+) contributions? on/g;
  for (const match of html.matchAll(countPattern)) {
    if (days.has(match[1])) days.get(match[1]).count = Number(match[2]);
  }

  const records = [...days.values()].sort((a, b) => a.date.localeCompare(b.date));
  if (!records.length) throw new Error("GitHub contribution calendar returned no days");

  const total = records.reduce((sum, day) => sum + day.count, 0);
  let longest = 0;
  let run = 0;
  for (const day of records) {
    run = day.count > 0 ? run + 1 : 0;
    longest = Math.max(longest, run);
  }

  let cursor = records.length - 1;
  if (records[cursor].count === 0) cursor -= 1;
  let current = 0;
  while (cursor >= 0 && records[cursor].count > 0) {
    current += 1;
    cursor -= 1;
  }

  return { total, current, longest, updated: isoDate(end) };
}

function overviewCard({ repositories, stars, followers, createdYear, updated }) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="420" height="170" viewBox="0 0 420 170" role="img" aria-labelledby="title desc">
  <title id="title">Abhishek's live GitHub overview</title>
  <desc id="desc">${repositories} public repositories, ${stars} stars, ${followers} followers. Automatically updated ${updated}.</desc>
  <defs>
    <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#F4F4F4"/><stop offset=".55" stop-color="#B2DCE2" stop-opacity=".62"/><stop offset="1" stop-color="#EDB2CE" stop-opacity=".72"/></linearGradient>
    <filter id="shadow" x="-10%" y="-15%" width="120%" height="130%"><feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#4A3E57" flood-opacity=".16"/></filter>
  </defs>
  <rect x="8" y="8" width="404" height="154" rx="20" fill="url(#bg)" stroke="#B39DE2" stroke-width="1.5" filter="url(#shadow)"/>
  <circle cx="44" cy="43" r="18" fill="#B39DE2"/><path d="M36 43h16M44 35v16" stroke="#2D2337" stroke-width="2.5" stroke-linecap="round"/>
  <text x="72" y="39" fill="#2D2337" font-family="Segoe UI,Arial,sans-serif" font-size="18" font-weight="700">GitHub overview</text>
  <text x="72" y="58" fill="#746A7D" font-family="Segoe UI,Arial,sans-serif" font-size="11">@${xml(username)} · automatically updated ${updated}</text>
  <g font-family="Segoe UI,Arial,sans-serif" text-anchor="middle">
    <text x="65" y="107" fill="#2D2337" font-size="25" font-weight="700">${repositories}</text><text x="65" y="126" fill="#4A3E57" font-size="11">REPOSITORIES</text>
    <text x="163" y="107" fill="#2D2337" font-size="25" font-weight="700">${stars}</text><text x="163" y="126" fill="#4A3E57" font-size="11">STARS</text>
    <text x="257" y="107" fill="#2D2337" font-size="25" font-weight="700">${followers}</text><text x="257" y="126" fill="#4A3E57" font-size="11">FOLLOWERS</text>
    <text x="355" y="107" fill="#2D2337" font-size="25" font-weight="700">${createdYear}</text><text x="355" y="126" fill="#4A3E57" font-size="11">SINCE</text>
  </g>
</svg>`;
}

function languageCard(languages, updated) {
  const palette = ["#B39DE2", "#B2DCE2", "#EDB2CE"];
  const top = languages.slice(0, 3);
  const max = Math.max(...top.map((item) => item.count), 1);
  const rows = top.map((item, index) => {
    const y = 91 + index * 28;
    const width = Math.round((item.count / max) * 269);
    return `<text x="30" y="${y}" fill="#2D2337" font-size="12" font-weight="600">${xml(item.name)}</text><rect x="111" y="${y - 11}" width="269" height="13" rx="6.5" fill="#FFFFFF" fill-opacity=".72"/><rect x="111" y="${y - 11}" width="${width}" height="13" rx="6.5" fill="${palette[index]}"/>`;
  }).join("");

  return `<svg xmlns="http://www.w3.org/2000/svg" width="420" height="170" viewBox="0 0 420 170" role="img" aria-labelledby="title desc">
  <title id="title">Abhishek's live language focus</title><desc id="desc">Primary languages across public repositories. Automatically updated ${updated}.</desc>
  <defs><linearGradient id="bg2" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#F4F4F4"/><stop offset=".5" stop-color="#F1B8D6" stop-opacity=".58"/><stop offset="1" stop-color="#B39DE2" stop-opacity=".58"/></linearGradient><filter id="shadow2" x="-10%" y="-15%" width="120%" height="130%"><feDropShadow dx="0" dy="4" stdDeviation="6" flood-color="#4A3E57" flood-opacity=".16"/></filter></defs>
  <rect x="8" y="8" width="404" height="154" rx="20" fill="url(#bg2)" stroke="#EDB2CE" stroke-width="1.5" filter="url(#shadow2)"/>
  <circle cx="44" cy="43" r="18" fill="#B2DCE2"/><path d="m36 43 6 6 11-13" fill="none" stroke="#2D2337" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
  <text x="72" y="39" fill="#2D2337" font-family="Segoe UI,Arial,sans-serif" font-size="18" font-weight="700">Language focus</text><text x="72" y="58" fill="#746A7D" font-family="Segoe UI,Arial,sans-serif" font-size="11">Repository languages · updated ${updated}</text>
  <g font-family="Segoe UI,Arial,sans-serif">${rows}</g>
</svg>`;
}

function streakCard({ total, current, longest, updated }) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="860" height="190" viewBox="0 0 860 190" role="img" aria-labelledby="title desc">
  <title id="title">Abhishek's live GitHub contribution streak</title><desc id="desc">${total} contributions in the trailing year, a ${current} day current streak, and a ${longest} day longest streak. Automatically updated ${updated}.</desc>
  <defs><linearGradient id="streakBg" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#B39DE2" stop-opacity=".42"/><stop offset=".35" stop-color="#B2DCE2" stop-opacity=".52"/><stop offset=".68" stop-color="#EDB2CE" stop-opacity=".52"/><stop offset="1" stop-color="#F1B8D6" stop-opacity=".5"/></linearGradient><linearGradient id="ring" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#B39DE2"/><stop offset="1" stop-color="#EDB2CE"/></linearGradient></defs>
  <rect x="8" y="8" width="844" height="174" rx="24" fill="#F4F4F4" stroke="#B39DE2" stroke-width="1.5"/><rect x="9" y="9" width="842" height="172" rx="23" fill="url(#streakBg)"/>
  <text x="34" y="41" fill="#2D2337" font-family="Segoe UI,Arial,sans-serif" font-size="19" font-weight="700">Contribution streak</text><text x="34" y="61" fill="#746A7D" font-family="Segoe UI,Arial,sans-serif" font-size="11">Trailing-year public activity · updated ${updated}</text>
  <line x1="286" y1="79" x2="286" y2="153" stroke="#FFFFFF" stroke-opacity=".72"/><line x1="574" y1="79" x2="574" y2="153" stroke="#FFFFFF" stroke-opacity=".72"/>
  <g font-family="Segoe UI,Arial,sans-serif" text-anchor="middle"><text x="145" y="116" fill="#2D2337" font-size="35" font-weight="700">${total}</text><text x="145" y="140" fill="#4A3E57" font-size="12" font-weight="600">YEARLY CONTRIBUTIONS</text><circle cx="430" cy="112" r="42" fill="#FFFFFF" fill-opacity=".68" stroke="url(#ring)" stroke-width="8"/><text x="430" y="120" fill="#2D2337" font-size="34" font-weight="700">${current}</text><text x="430" y="171" fill="#7B5DB3" font-size="12" font-weight="700">CURRENT STREAK · DAYS</text><text x="715" y="116" fill="#2D2337" font-size="35" font-weight="700">${longest}</text><text x="715" y="140" fill="#4A3E57" font-size="12" font-weight="600">LONGEST STREAK · DAYS</text></g>
  <path d="M430 63c-8-10-1-19 6-24-1 8 7 10 7 18 0 7-5 12-13 12-8 0-13-5-13-12 0-6 3-10 8-14-1 7 2 11 5 12" fill="#EDB2CE" stroke="#7B5DB3" stroke-width="2" stroke-linejoin="round"/>
</svg>`;
}

const user = await request(`https://api.github.com/users/${username}`).then((response) => response.json());
const repositories = await getRepositories();
const updated = isoDate(new Date());
const stars = repositories.reduce((sum, repository) => sum + repository.stargazers_count, 0);
const languageCounts = new Map();
for (const repository of repositories) {
  if (repository.language) {
    languageCounts.set(repository.language, (languageCounts.get(repository.language) || 0) + 1);
  }
}
const languages = [...languageCounts.entries()]
  .map(([name, count]) => ({ name, count }))
  .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
const streak = await getContributionStats();

await mkdir("assets", { recursive: true });
await Promise.all([
  writeFile("assets/github-overview.svg", overviewCard({ repositories: repositories.length, stars, followers: user.followers, createdYear: new Date(user.created_at).getUTCFullYear(), updated })),
  writeFile("assets/language-focus.svg", languageCard(languages, updated)),
  writeFile("assets/github-streak.svg", streakCard(streak)),
]);

console.log(JSON.stringify({ username, repositories: repositories.length, stars, followers: user.followers, languages, streak }, null, 2));
