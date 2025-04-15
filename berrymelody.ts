import * as xlsx from "xlsx";
import { WikiWiki } from "./wikiwiki";

const WIKI_NAME = "berrymelody";
const PASSWORD = process.env.PASSWORD || "password";
const EXCEPTIONS: {
  before: string;
  after: string;
}[] = [];
const IGNORED_PAGES: string[] = [];
const CHAPTERS = {
  "単曲 / Ayira's Suitcase": "single",
  "序章 / Da Capo": "chap0",
  "章节一 / Sotto voce": "chap1",
  "章节二 / Variazione": "chap2",
  "章节三 / Fortepiano": "chap3",
  "联动集 / Notanote": "notanote",
  "炫光动感 / Dynamix": "dynamix",
  "联动集 / DanceRail3": "dancerail3",
  "联动集 / Phigros": "phigros",
  "联动集 / Orzmic": "orzmic",
  "求闻音录 / THMR": "thmr",
  "调律诗篇 / Lanota": "lanota",
  "火花线 / SparkLine": "sparkline",
  "喵赛克 / MUSYNC": "musync",
  "卉 / HUI-Works": "hui-works",
  "精选集 / PYKAMIA": "pykamia",
} as {
  [key: string]: string;
};

const COLORS = {
  REALITY: "deepskyblue",
  ILLUSION: "orange",
  TWIST: "red",
  DREAMY: "mediumspringgreen",
  RUIN: "purple",
  FOOL: "lime",
} as {
  [key: string]: string;
};

const ABBREVIATIONS = {
  RL: "REALITY",
  IL: "ILLUSION",
  TT: "TWIST",
  DM: "DREAMY",
  RU: "RUIN",
  FL: "FOOL",
} as {
  [key: string]: string;
};

function readExcel(filePath: string): any[] {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet, { range: 4 });
}

function generateDifficulties(data: any): string {
  const start = data["RL"]
    ? "RL"
    : data["IL"]
    ? "IL"
    : data["TT"]
    ? "TT"
    : data["DM"]
    ? "DM"
    : "RU";
  const difficulties = ["RL", "IL", "TT", "DM", "RU"]
    .map((difficulty: string) => {
      if (data[difficulty]) {
        if (difficulty === start) {
          return `|~Difficulty|~&color(${COLORS[ABBREVIATIONS[difficulty]]}){${
            ABBREVIATIONS[difficulty]
          }};|${data[`${difficulty}_1`]}|~Notes|${data[`${difficulty}_2`]}|`;
        }
        return `|~|~&color(${COLORS[ABBREVIATIONS[difficulty]]}){${
          ABBREVIATIONS[difficulty]
        }};|${data[`${difficulty}_1`]}|~|${data[`${difficulty}_2`]}|`;
      } else {
        return "";
      }
    })
    .filter((difficulty) => difficulty !== "");
  return difficulties.join("\n");
}

function generateCharters(data: any): string {
  // もし楽曲に存在する難易度のCharterがすべて同じであれば、まとめる
  const charter = ["RL", "IL", "TT", "DM", "RU"]
    .map((difficulty: string) => {
      if (data[difficulty]) {
        return data[difficulty];
      } else {
        return "";
      }
    })
    .filter((charter) => charter !== "");
  if (charter.every((c) => c === charter[0])) {
    return `|~Charter|>|>|>|${charter[0]}|`;
  }

  const start = data["RL"]
    ? "RL"
    : data["IL"]
    ? "IL"
    : data["TT"]
    ? "TT"
    : data["DM"]
    ? "DM"
    : "RU";
  const charters = ["RL", "IL", "TT", "DM", "RU"]
    .map((difficulty: string) => {
      if (data[difficulty]) {
        if (difficulty === start) {
          return `|~Charter|~&color(${COLORS[ABBREVIATIONS[difficulty]]}){${
            ABBREVIATIONS[difficulty]
          }};|>|>|${data[difficulty]}|`;
        }
        return `|~|~&color(${COLORS[ABBREVIATIONS[difficulty]]}){${
          ABBREVIATIONS[difficulty]
        }};|>|>|${data[difficulty]}|`;
      } else {
        return "";
      }
    })
    .filter((charter) => charter !== "");
  return charters.join("\n");
}

function generateWikiContent(data: any): string {
  // PukiWiki形式の内容を生成する
  return `|CENTER:150|CENTER:130|CENTER:130|CENTER:130|CENTER:130|c
|>|>|>|>|&attachref();|
|~Composer|>|>|>|[[${data["コンポーザー"]}>アーティスト順#]]|
${generateCharters(data)}
|~Artwork|>|>|>|${data["アートワーク"]}|
${generateDifficulties(data)}
|~Length|>|>|>|${data["曲の長さ"]}|
|~BPM|>|>|>|${data["BPM"]}|
|~Chapter|>|>|>|[[${data["チャプター"]}>チャプター順#${
    CHAPTERS[data["チャプター"]]
  }]]|

* 解説 [#Guide]
${data["RL"] ? `*** &color(deepskyblue){REALITY}; [#RL]` : ""}${
    data["IL"] ? `\n*** &color(orange){ILLUSION}; [#IL]` : ""
  }${data["TT"] ? `\n*** &color(red){TWIST}; [#TT]` : ""}${
    data["DM"] ? `\n*** &color(mediumspringgreen){DREAMY}; [#DM]` : ""
  }${data["RU"] ? `\n*** &color(purple){RUIN}; [#RU]` : ""}

* 音源 [#Audio]
*** YouTube [#YouTube]
//#youtube()
*** SoundCloud [#SoundCloud]
//#soundcloud()

* プレイ動画 [#Play]

* コメント [#Comment]
#pcomment(,10,noname,above,below,reply)
`;
}

function updateWikiContent(source: string, data: any): string {
  const lines = source.split("\n");
  const difficultyRegex =
    /\|~\|~&color\([A-Za-z]+\)\{[A-Za-z]+\};\|[0-9+]+\|~\|.+\|/;
  const updatedLines = lines.map((line) => {
    if (line.startsWith("|~Composer|")) {
      return line;
    } else if (line.startsWith("|~Artwork|")) {
      return `|~Artwork|>|>|>|${data["アートワーク"]}|`;
    } else if (line.startsWith("|~Difficulty|")) {
      return generateDifficulties(data);
    } else if (line.startsWith("|~BPM|")) {
      if (CHAPTERS[data["チャプター"]] === "single") {
        return line;
      }
      return `|~BPM|>|>|>|${data["BPM"]}|`;
    } else if (line.startsWith("|~Length|")) {
      return `|~Length|>|>|>|${data["曲の長さ"]}|`;
    } else if (line.startsWith("|~Chapter|")) {
      return `|~Chapter|>|>|>|[[${data["チャプター"]}>チャプター順#${
        CHAPTERS[data["チャプター"]]
      }]]|`;
    } else if (difficultyRegex.test(line)) {
      return undefined;
    } else {
      return line;
    }
  });
  return updatedLines.filter((line) => line !== undefined).join("\n");
}

async function main() {
  try {
    const wiki = new WikiWiki(WIKI_NAME, PASSWORD);
    await wiki.authenticate();
    console.log("Authenticated successfully.");

    const existingPages = await wiki.getPages();
    console.log("Fetched existing pages.");

    const excelData = readExcel("berrymelody.xlsx");
    console.log("Loaded Excel data.");

    const targetChapters = ["単曲 / Ayira's Suitcase"];

    for (const row of excelData) {
      const replacedName: string = EXCEPTIONS.reduce(
        (acc, exception) => acc.replace(exception.before, exception.after),
        row["曲名"]
      );
      if (IGNORED_PAGES.includes(replacedName)) {
        continue;
      }
      // 誤爆防止
      if (!targetChapters.includes(row["チャプター"])) {
        continue;
      }
      if (existingPages.includes(replacedName)) {
        const source = await wiki.getPage(replacedName);
        const updatedSource = updateWikiContent(source, row);
        await wiki.updatePage(replacedName, updatedSource);
        console.log(`Updated page: ${replacedName}`);
      } else {
        const source = generateWikiContent(row);
        await wiki.updatePage(replacedName, source);
        console.log(`Created page: ${replacedName}`);
      }
    }

    console.log("Wiki pages updated successfully.");
  } catch (error) {
    if (error instanceof Error) {
      console.error(error.message);
    } else {
      console.error("An unknown error occurred.");
    }
  }
}

main();
