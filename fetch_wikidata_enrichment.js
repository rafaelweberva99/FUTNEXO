const fs = require("fs");
const vm = require("vm");

const ROOT = __dirname;
const INDEX_PATH = ROOT + "\\index.html";
const OUTPUT_PATH = ROOT + "\\futbrain_enrichment.json";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchJson(url) {
  const res = await fetch(url, {
    headers: {
      accept: "application/sparql-results+json, application/json",
      "user-agent": "FutBrainDataBot/1.0 (local project maintenance)",
    },
  });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}: ${await res.text()}`);
  }
  return await res.json();
}

function escapeSparqlString(value) {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function yearFromIso(value) {
  if (!value) return null;
  const match = String(value).match(/^(\d{4})/);
  return match ? Number(match[1]) : null;
}

function extractDb() {
  const text = fs.readFileSync(INDEX_PATH, "utf8");
  const start = text.indexOf("const DB = {");
  const end = text.indexOf("const TRIVIA=[");
  if (start === -1 || end === -1) {
    throw new Error("No se pudo encontrar el bloque DB en index.html");
  }
  const context = { console, globalThis: {} };
  vm.createContext(context);
  vm.runInContext(text.slice(start, end) + "\n;globalThis.__DB__ = DB;", context);
  return context.globalThis.__DB__;
}

function uniqueClubRows(rows) {
  const seen = new Set();
  return rows
    .filter((row) => row.club_name)
    .filter((row) => {
      const key = `${row.club_name}|${row.from_year ?? ""}|${row.to_year ?? ""}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .sort((a, b) => {
      const ay = a.from_year ?? 9999;
      const by = b.from_year ?? 9999;
      if (ay !== by) return ay - by;
      const aty = a.to_year ?? 9999;
      const bty = b.to_year ?? 9999;
      if (aty !== bty) return aty - bty;
      return a.club_name.localeCompare(b.club_name);
    });
}

function cleanPhotoUrl(value) {
  if (!value) return "";
  const text = String(value);
  if (text.startsWith("http://") || text.startsWith("https://")) {
    if (text.includes("Special:FilePath/http")) {
      const match = text.match(/Special:FilePath\/(.*)$/);
      if (match && match[1]) {
        try {
          return decodeURIComponent(match[1]);
        } catch {
          return match[1];
        }
      }
    }
    return text;
  }
  return "https://commons.wikimedia.org/wiki/Special:FilePath/" + encodeURIComponent(text);
}

function isClubLike(clubName) {
  const text = String(clubName || "").toLowerCase();
  const flat = text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
  if (!text) return false;
  if (text.includes("national team")) return false;
  if (text.includes("national football team")) return false;
  if (text.includes("national association football team")) return false;
  if (flat.includes("mens national")) return false;
  if (flat.includes("womens national")) return false;
  if (flat.includes("men s national")) return false;
  if (flat.includes("women s national")) return false;
  if (text.includes("olympic")) return false;
  if (text.includes("under-")) return false;
  if (text.includes("u-")) return false;
  if (/\bsub[\s-]?\d{1,2}\b/.test(text)) return false;
  if (/\bu[\s-]?\d{1,2}\b/.test(text)) return false;
  if (/\bunder[\s-]?\d{1,2}\b/.test(text)) return false;
  if (/\breserve(s)?\b/.test(text)) return false;
  if (/\b(reserva|reservas)\b/.test(text)) return false;
  if (/\bjuv(enil|eniles)?\b/.test(text)) return false;
  if (/\byouth\b/.test(text)) return false;
  if (/\bprimavera\b/.test(text)) return false;
  if (/\bbarcelona b\b/.test(text)) return false;
  if (/\bbarcelona c\b/.test(text)) return false;
  if (flat.includes("barcelona c")) return false;
  if (flat.includes("barcelona atletic")) return false;
  if (/\breal madrid castilla\b/.test(text)) return false;
  if (/\breal madrid c\b/.test(text)) return false;
  if (/\batletico madrid b\b/.test(text)) return false;
  if (/\bsevilla atletico\b/.test(text)) return false;
  if (/\bjuventus next gen\b/.test(flat)) return false;
  if (/\b[a-z0-9 .'-]+ b\b/.test(text)) return false;
  if (/\b[a-z0-9 .'-]+ c\b/.test(text)) return false;
  if (/\b[a-z0-9 .'-]+ ii\b/.test(text)) return false;
  return true;
}

function normalizeClubName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\b(futbol|futebol|football)\b/g, " ")
    .replace(/\b(club|clube|clube de regatas|club social y deportivo|sporting club|sociedad anonima deportiva)\b/g, " ")
    .replace(/\b(fc|cf|ac|as|sc|cd|ca)\b/g, " ")
    .replace(/\b(1913|1905|1903|1900)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function canonicalClubName(value) {
  const norm = normalizeClubName(value);
  const aliases = [
    ["Barcelona", ["barcelona"]],
    ["Real Madrid", ["real madrid", "real madrid club de futbol"]],
    ["Atlético Madrid", ["atletico madrid"]],
    ["Southampton", ["southampton", "southampton f c", "southampton fc"]],
    ["Bayern Munich", ["bayern munich", "bayern munchen", "fc bayern munich", "fc bayern munchen"]],
    ["Borussia Dortmund", ["borussia dortmund"]],
    ["Inter Milan", ["inter milan", "internazionale", "fc internazionale milano"]],
    ["Milan", ["milan", "ac milan"]],
    ["Juventus", ["juventus", "juventus fc"]],
    ["PSG", ["psg", "paris saint germain"]],
    ["Lyon", ["lyon", "olympique lyonnais"]],
    ["Manchester United", ["manchester united", "manchester united f c", "manchester united fc"]],
    ["Manchester City", ["manchester city"]],
    ["Chelsea", ["chelsea"]],
    ["Liverpool", ["liverpool"]],
    ["Arsenal", ["arsenal"]],
    ["Tottenham", ["tottenham", "tottenham hotspur", "tottenham hotspur f c", "tottenham hotspur fc"]],
    ["Sevilla", ["sevilla"]],
    ["Valencia", ["valencia"]],
    ["Roma", ["roma", "as roma"]],
    ["Lazio", ["lazio", "ss lazio"]],
    ["Napoli", ["napoli", "ssc napoli"]],
    ["Porto", ["porto", "fc porto"]],
    ["Benfica", ["benfica", "sl benfica"]],
    ["Ajax", ["ajax", "afc ajax"]],
    ["Monaco", ["monaco", "as monaco"]],
    ["Sporting CP", ["sporting cp", "sporting clube de portugal"]],
    ["LAFC", ["los angeles fc", "los angeles football club", "lafc"]],
    ["Al Nassr", ["al nassr", "al nassr fc", "al-nassr"]],
    ["Al Hilal", ["al hilal", "al hilal sfc"]],
    ["Al Ittihad", ["al ittihad", "al ittihad fc"]],
    ["Athletico Paranaense", ["athletico paranaense", "club athletico paranaense"]],
    ["Flamengo", ["flamengo", "clube de regatas do flamengo"]],
    ["Colo-Colo", ["colo colo", "colocolo", "club social y deportivo colo colo"]],
    ["Universidad de Chile", ["universidad de chile", "club universidad de chile"]],
    ["Boca Juniors", ["boca juniors", "club atletico boca juniors"]],
    ["River Plate", ["river plate", "club atletico river plate"]],
    ["Santos", ["santos", "santos fc"]],
    ["Palmeiras", ["palmeiras", "se palmeiras"]],
    ["São Paulo", ["sao paulo", "sao paulo fc"]],
    ["Fluminense", ["fluminense", "fluminense fc"]],
    ["Parma", ["parma", "parma calcio 1913"]],
    ["Inter Miami", ["inter miami", "club internacional de futbol miami"]],
    ["Newell's Old Boys", ["newell s old boys", "club atletico newell s old boys"]],
  ];
  for (const [canonical, variants] of aliases) {
    if (variants.includes(norm)) return canonical;
  }
  return value;
}

async function queryBatch(batch) {
  const values = batch
    .map((item) => `"${escapeSparqlString(item.query_name)}"@en`)
    .join(" ");

  const query = `
SELECT ?queryName ?player ?birth ?countryLabel ?imageUrl ?clubLabel ?start ?end WHERE {
  VALUES ?queryName { ${values} }
  ?player rdfs:label ?queryName.
  ?player wdt:P31 wd:Q5.
  ?player wdt:P106 wd:Q937857.
  OPTIONAL { ?player wdt:P569 ?birth. }
  OPTIONAL { ?player wdt:P27 ?country. }
  OPTIONAL {
    ?player wdt:P18 ?image.
    BIND(CONCAT("https://commons.wikimedia.org/wiki/Special:FilePath/", ENCODE_FOR_URI(STR(?image))) AS ?imageUrl)
  }
  OPTIONAL {
    ?player p:P54 ?stmt.
    ?stmt ps:P54 ?club.
    OPTIONAL { ?stmt pq:P580 ?start. }
    OPTIONAL { ?stmt pq:P582 ?end. }
  }
  SERVICE wikibase:label { bd:serviceParam wikibase:language "en". }
}`.trim();

  const url =
    "https://query.wikidata.org/sparql?format=json&query=" + encodeURIComponent(query);
  return await fetchJson(url);
}

async function main() {
  const DB = extractDb();
  const players = Object.keys(DB)
    .sort()
    .map((name, idx) => ({
      player_id: `P${String(idx + 1).padStart(4, "0")}`,
      name,
      query_name: DB[name].wiki || name,
      flag: DB[name].flag || "",
      current_clubs: (DB[name].clubs || []).map((club) => ({
        club_name: club.n || "",
        from_year: club.f ?? null,
        to_year: club.t ?? null,
      })),
    }));

  const grouped = new Map();
  for (const player of players) {
    grouped.set(player.query_name, {
      player_id: player.player_id,
      name: player.name,
      query_name: player.query_name,
      flag: player.flag,
      current_clubs: player.current_clubs,
      wikipedia_url: `https://en.wikipedia.org/wiki/${player.query_name.replace(/ /g, "_")}`,
      wikidata_url: "",
      photo_url: "",
      birth_year: null,
      country: "",
      wikidata_clubs: [],
      status: "sin_datos",
      notes: "",
      candidates: {},
    });
  }

  const batches = [];
  for (let i = 0; i < players.length; i += 20) {
    batches.push(players.slice(i, i + 20));
  }

  for (let i = 0; i < batches.length; i++) {
    const batch = batches[i];
    const data = await queryBatch(batch);
    const bindings = data?.results?.bindings || [];

    for (const row of bindings) {
      const queryName = row.queryName?.value || "";
      const item = grouped.get(queryName);
      if (!item) continue;
      const playerUri = row.player?.value || "";
      if (!playerUri) continue;
      if (!item.candidates[playerUri]) {
        item.candidates[playerUri] = {
          wikidata_url: playerUri,
          photo_url: "",
          birth_year: null,
          country: "",
          wikidata_clubs: [],
        };
      }
      const candidate = item.candidates[playerUri];
      if (row.imageUrl?.value && !candidate.photo_url) candidate.photo_url = cleanPhotoUrl(row.imageUrl.value);
      if (row.birth?.value && !candidate.birth_year) candidate.birth_year = yearFromIso(row.birth.value);
      if (row.countryLabel?.value && !candidate.country) candidate.country = row.countryLabel.value;
      if (row.clubLabel?.value && isClubLike(row.clubLabel.value)) {
        candidate.wikidata_clubs.push({
          club_name: canonicalClubName(row.clubLabel.value),
          from_year: yearFromIso(row.start?.value || null),
          to_year: yearFromIso(row.end?.value || null),
        });
      }
    }

    console.log(`Batch ${i + 1}/${batches.length} listo`);
    await sleep(1200);
  }

  const result = {
    generated_at: new Date().toISOString(),
    players: [],
  };

  for (const player of players) {
    const item = grouped.get(player.query_name);
    const currentNorm = item.current_clubs.map((club) => normalizeClubName(club.club_name));

    let bestCandidate = null;
    let bestScore = -1;
    for (const candidate of Object.values(item.candidates)) {
      candidate.wikidata_clubs = uniqueClubRows(candidate.wikidata_clubs);
      const candidateNorm = candidate.wikidata_clubs.map((club) => normalizeClubName(club.club_name));
      const overlap = currentNorm.filter((club) => club && candidateNorm.includes(club)).length;
      const score = overlap * 100 + candidate.wikidata_clubs.length * 3 + (candidate.country ? 5 : 0) + (candidate.photo_url ? 2 : 0);
      if (score > bestScore) {
        bestScore = score;
        bestCandidate = candidate;
      }
    }

    if (bestCandidate) {
      item.wikidata_url = bestCandidate.wikidata_url;
      item.photo_url = bestCandidate.photo_url;
      item.birth_year = bestCandidate.birth_year;
      item.country = bestCandidate.country;
      item.wikidata_clubs = bestCandidate.wikidata_clubs;
    } else {
      item.wikidata_clubs = [];
    }

    const currentComp = item.current_clubs.map(
      (club) => `${club.club_name} (${club.from_year ?? "?"}-${club.to_year ?? "?"})`
    );
    const wdComp = item.wikidata_clubs.map(
      (club) => `${club.club_name} (${club.from_year ?? "?"}-${club.to_year ?? "?"})`
    );

    const wikiNorm = item.wikidata_clubs.map((club) => normalizeClubName(club.club_name));
    const missingCurrentClubs = currentNorm.filter((club) => club && !wikiNorm.includes(club));

    if (!item.wikidata_url) {
      item.status = "sin_wikidata";
      item.notes = "No se encontró coincidencia exacta en Wikidata.";
    } else if (!item.wikidata_clubs.length) {
      item.status = "sin_clubes";
      item.notes = "Wikidata no devolvió historial de clubes.";
    } else if (missingCurrentClubs.length) {
      item.status = "revisar";
      item.notes = "Faltan clubes de la base actual dentro de la trayectoria devuelta por Wikidata.";
    } else if (JSON.stringify(currentComp) !== JSON.stringify(wdComp)) {
      item.status = "ampliado";
      item.notes = "Wikidata aporta más tramos o temporadas que la base actual.";
    } else {
      item.status = "ok";
      item.notes = "";
    }

    result.players.push(item);
  }

  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(result, null, 2), "utf8");
  const reviewCount = result.players.filter((player) => player.status !== "ok").length;
  console.log(`Guardado: ${OUTPUT_PATH}`);
  console.log(`Jugadores: ${result.players.length}`);
  console.log(`Para revisar: ${reviewCount}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
