import * as xlsx from "xlsx";
import { WikiWiki } from "./wikiwiki-ext";

const WIKI_NAME = "orzmic";
const EXCEPTIONS: {
  before: string;
  after: string;
}[] = [
  {
    before: "UИ0HЯ1Λ: Seventh Syberia Strike",
    after: "UN0HR1A",
  },
  {
    before: "Przepływ światła",
    after: "Przeplyw swiatla",
  },
  {
    before: "雨霖铃",
    after: "雨霖鈴",
  },
  {
    before: "SCHADEN ~Schubert: Erlkönig, D.328~",
    after: "SCHADEN",
  },
  {
    before: "Äventyr",
    after: "Aventyr",
  },
  {
    before: "遗世之地",
    after: "遺世之地",
  },
  {
    before: "Proto:Messiah",
    after: "Proto Messiah",
  },
  {
    before: "心機一轉[白]",
    after: "心機一轉(白)",
  },
  {
    before: "エセ交響詩[若き白の女王]",
    after: "エセ交響詩(若き白の女王)",
  },
];
const IGNORED_PAGES: string[] = [];
const CHAPTERS = {
  "Chapter 1~3 / Era of Opera": "chap1",
  "Chapter 4 / Unexpected": "chap4",
  "Chapter 5 / Mystery Puppet": "chap5",
  "Chapter 6 / THE MAZE": "chap6",
  "Chapter 7 / Evernight Curse": "chap7",
  "Chapter 8 / Devil's Chessboard": "chap8",
  "Chapter 9 / Afterglow": "chap9",
  "Chapter 0_1 / Single Songs": "chap0-1",
  "Chapter 0_2 / Single Songs": "chap0-2",
  "Chapter EX / Song Package 1": "exone",
  "Chapter EX / Song Package 2": "extwo",
  "Chapter EX / Song Package 3": "exthree",
  "Chapter EX / Song Package 4": "exfour",
  "Chapter EX / Song Package 5": "exfive",
  "Chapter SP / Broken Mirage": "mirage",
  "Chapter EX / Flower and Planet": "flower",
  "Chapter EX / special collaboration": "special",
  "Chapter CO1 / Dance Cube Collaboration": "co1",
  "Chapter CO2 / MUSYNC Collaboration": "co2",
  "Chapter CO3 / SparkLine Collaboration": "co3",
  "Chapter CO4 / Lanota Collaboration": "co4",
  "Chapter CO5 / GTS Collaboration": "co5",
  "Chapter ? / Song Package ?": "april141",
} as {
  [key: string]: string;
};

const COLORS = {
  EASY: "deepskyblue",
  NORMAL: "orange",
  HARD: "red",
  SPECIAL: "black",
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
  const start = data["EASY"]
    ? "EASY"
    : data["NORMAL"]
    ? "NORMAL"
    : data["HARD"]
    ? "HARD"
    : "SPECIAL";
  const difficulties = ["EASY", "NORMAL", "HARD", "SPECIAL"]
    .map((difficulty: string) => {
      if (data[difficulty]) {
        if (difficulty === start) {
          return `|~Difficulty|~&color(${COLORS[difficulty]}){${difficulty}};|${
            data[`${difficulty}_1`]
          }|~Rating|${Number(data[`${difficulty}_2`]).toFixed(1)}|~Notes|${
            data[`${difficulty}_3`]
          }|`;
        }
        return `|~|~&color(${COLORS[difficulty]}){${difficulty}};|${
          data[`${difficulty}_1`]
        }|~|${Number(data[`${difficulty}_2`]).toFixed(1)}|~|${
          data[`${difficulty}_3`]
        }|`;
      } else {
        return "";
      }
    })
    .filter((difficulty) => difficulty !== "");
  return difficulties.join("\n");
}

function generateCharters(data: any): string {
  const COLORS: Record<string, string> = {
    EASY: "deepskyblue",
    NORMAL: "orange",
    HARD: "red",
    SPECIAL: "black",
  };

  // 制作者ごとに難易度をまとめる
  const charterMap: Record<string, string[]> = {};

  ["EASY", "NORMAL", "HARD", "SPECIAL"].forEach((difficulty) => {
    if (data[difficulty]) {
      const charter = data[difficulty];
      if (!charterMap[charter]) {
        charterMap[charter] = [];
      }
      charterMap[charter].push(difficulty);
    }
  });

  // すべての難易度で同じ制作者なら1行にまとめる
  const charters = Object.keys(charterMap);
  if (charters.length === 1) {
    return `|~Chart Designer|>|>|>|>|>|${charters[0]}|`;
  }

  // 制作者ごとに出力
  return charters
    .map((charter, index) => {
      const difficulties = charterMap[charter]
        .map((diff) => `&color(${COLORS[diff]}){${diff}};`)
        .join(" / ");

      if (index === 0) {
        return `|~Chart Designer|~${difficulties}|>|>|>|>|${charter}|`;
      } else {
        return `|~|~${difficulties}|>|>|>|>|${charter}|`;
      }
    })
    .join("\n");
}

function generateWikiContent(data: any): string {
  // PukiWiki形式の内容を生成する
  return `|CENTER:150|CENTER:130|CENTER:130|CENTER:130|CENTER:130|CENTER:130|CENTER:130|c
|>|>|>|>|>|>|&attachref();|
|~Composer|>|>|>|>|>|[[${data["コンポーザー"]}>アーティスト順#]]|
${generateCharters(data)}
|~Cover Artist|>|>|>|>|>|${data["カバーアーティスト（絵）"]}|
${generateDifficulties(data)}
|~Length|>|>|>|>|>|${data["曲の長さ"] || "*:**"}|
|~BPM|>|>|>|>|>|${data["BPM"] || "***"}|
${
  ["chap0-1", "chap0-2"].includes(CHAPTERS[data["チャプター"]])
    ? `|~Chapter|>|>|>|>|>|[[${data["チャプター"]}>楽曲一覧#${
        CHAPTERS[data["チャプター"]]
      }]]|`
    : `|~Chapter|>|>|>|[[${data["チャプター"]}>楽曲一覧#${
        CHAPTERS[data["チャプター"]]
      }]]|~Song Number|*|`
}

* 解説 [#Guide]
${data["EASY"] ? `*** EASY [#EASY]` : ""}${
    data["NORMAL"] ? `\n*** NORMAL [#NORMAL]` : ""
  }${data["HARD"] ? `\n*** HARD [#HARD]` : ""}${
    data["SPECIAL"] ? `\n*** SPECIAL [#SPECIAL]` : ""
  }

* 音源 [#Audio]
*** YouTube [#YouTube]
//#youtube()
*** SoundCloud [#SoundCloud]
//#soundcloud()
* プレイ動画 [#Play]
// 例
// - 難易度：''&color(Red){HARD};''[1,000,000Pts]
// Player : プレイヤー名
// #youtube(動画ID)
// #br
* コメント [#Comment]
#pcomment(,10,noname,above,below,reply)
`;
}

function updateWikiContent(source: string, data: any): string {
  const lines = source.split("\n");
  const difficultyRegex =
    /\|~\|~&color\([A-Za-z]+\)\{[A-Za-z]+\};\|[0-9+]+\|~\|.+\|~\|.+\|/;
  const updatedLines = lines.map((line) => {
    if (line.startsWith("|~Composer|")) {
      return line;
    } else if (line.startsWith("|~Cover Artist|")) {
      return `|~Cover Artist|>|>|>|${data["カバーアーティスト（絵）"]}|`;
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
      return `|~Chapter|>|>|>|[[${data["チャプター"]}>楽曲一覧#${
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
    const wiki = new WikiWiki(WIKI_NAME);

    const existingPages = await wiki.getPages();
    console.log("Fetched existing pages.");

    const excelData = readExcel("orzmic.xlsx");
    console.log("Loaded Excel data.");

    const targetChapters = [
      "Chapter CO5 / GTS Collaboration",
      "Chapter ? / Song Package ?",
    ];

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
        //console.log(updatedSource);
      } else {
        const source = generateWikiContent(row);
        console.log(source);
      }
      console.log("--------------------");
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
