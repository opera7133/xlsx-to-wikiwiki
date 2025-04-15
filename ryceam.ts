import * as xlsx from "xlsx";
import { WikiWiki } from "./wikiwiki";

const WIKI_NAME = "ryceam";
const PASSWORD = process.env.PASSWORD || "password";
const EXCEPTIONS = [
  {
    before: "Spacecture: void",
    after: "Spacecture void",
  },
];
const IGNORED_PAGES = ["Microcosm Blowup", "+ERABY+E CONNEC+10N", "MIRЯOЯ"];
const CHAPTERS = {
  "Beginner / 薄明り": "Beginner",
  "Prelude / 鏡の間": "Prelude",
  "Chapter I / 鏡の源流": "Chapter_I",
  "Chapter II / 裂けた昼": "Chapter_II",
  "Chapter III / 花の追憶": "Chapter_III",
  "Subject Cosmos / 太初の夜": "Subject_Cosmos",
  "Subject Kawaii / レモン清夏": "Subject_Kawaii",
  "Subject Elec / 掣電一閃": "Subject_Elec",
  "Subject Spring / 重錦韶光": "Subject_Spring",
  "Subject Mix / シングル": "Subject_Mix",
  "Collaboration I / 次元LAB": "Dimension_LAB",
  "Collaboration II / DANCE CUBE": "DANCE_CUBE",
  "Collaboration III / MUSYNC": "MUSYNC",
  "Collaboration III / SONIC SURGE II": "SONIC_SURGE_II",
  "Legacy / 過去の章": "Legacy",
} as {
  [key: string]: string;
};

function readExcel(filePath: string): any[] {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet, { range: 4 });
}

function generateWikiContent(data: any): string {
  // PukiWiki形式の内容を生成する
  return `|CENTER:150|CENTER:100|CENTER:100|CENTER:100|CENTER:100|c
|>|>|>|>|&attachref();|
|>|>|>|>|LEFT:~&size(20){${data["曲名"]}};|
|~Composer|>|>|>|${data["コンポーザー"]}|
|~Artwork|>|>|>|${data["アートワーク"]}|
|~Difficulty|~&nbsp;|~Level|~Notes|~Charter|${
    data["Easy"]
      ? `\n|~|~&color(orange){EZ};|${data["Easy_1"]}|${data["Easy_2"]}|${data["Easy"]}|`
      : ""
  }${
    data["Parallel"]
      ? `\n|~|~&color(darkturquoise){PL};|${data["Parallel_1"]}|${
          data["Parallel_2"]
        }|${data["Easy"] === data["Parallel"] ? "~" : data["Parallel"]}|`
      : ""
  }${
    data["Rhythm"]
      ? `\n|~|~&color(dodgerblue){RY};|${data["Rhythm_1"]}|${
          data["Rhythm_2"]
        }|${data["Rhythm"] === data["Parallel"] ? "~" : data["Rhythm"]}|`
      : ""
  }${
    data["Havoc"]
      ? `\n|~|~&color(crimson){HC};|${data["Havoc_1"]}|${data["Havoc_2"]}|${
          data["Rhythm"] === data["Havoc"] ? "~" : data["Havoc"]
        }|`
      : ""
  }
|~BPM|>|${data["BPM"]}|~Length|${data["曲の長さ"]}|
|~Version|>|>|>|${data["追加"]}|
|~Unlock|~&color(darkturquoise){PL};|>|>|プリズム：2500|
|~|~&color(dodgerblue){RY};|>|>|~|
|~Chapter|>|>|>|[[${data["チャプター"]}>チャプター順#${
    CHAPTERS[data["チャプター"]]
  }]]|

* 楽曲概要 [#SongInfo]

* 攻略 [#Charts]
${data["Easy"] ? `** &color(orange){Easy}; [#Easy]` : ""}${
    data["Parallel"] ? `\n** &color(darkturquoise){Parallel}; [#Parallel]` : ""
  }${data["Rhythm"] ? `\n** &color(dodgerblue){Rhythm}; [#Rhythm]` : ""}${
    data["Havoc"] ? `\n** &color(crimson){Havoc}; [#Havoc]` : ""
  }

* 音源 [#Audio]
** YouTube [#YouTube]
//#youtube()
** SoundCloud [#SoundCloud]
//#soundcloud()

* プレイ動画 [#Play]

* コメント [#Comment]
#pcomment(,10,noname,above,below,reply)
`;
}

function updateWikiContent(source: string, data: any): string {
  const lines = source.split("\n");
  const difficultyRegex =
    /\|~\|~&color\([A-Za-z]+\)\{[A-Za-z]+\};\|[0-9]+\|[0-9]+\|.+\|/;
  const updatedLines = lines.map((line) => {
    if (line.startsWith("|~Composer|")) {
      return `|~Composer|>|>|>|${data["コンポーザー"]}|`;
    } else if (line.startsWith("|~Artwork|")) {
      return `|~Artwork|>|>|>|${data["アートワーク"]}|`;
    } else if (line.startsWith("|~Difficulty|")) {
      return `|~Difficulty|~&nbsp;|~Level|~Notes|~Charter|${
        data["Easy"]
          ? `\n|~|~&color(orange){EZ};|${data["Easy_1"]}|${data["Easy_2"]}|${data["Easy"]}|`
          : ""
      }${
        data["Parallel"]
          ? `\n|~|~&color(darkturquoise){PL};|${data["Parallel_1"]}|${
              data["Parallel_2"]
            }|${data["Parallel"] === data["Easy"] ? "~" : data["Parallel"]}|`
          : ""
      }${
        data["Rhythm"]
          ? `\n|~|~&color(dodgerblue){RY};|${data["Rhythm_1"]}|${
              data["Rhythm_2"]
            }|${data["Rhythm"] === data["Parallel"] ? "~" : data["Rhythm"]}|`
          : ""
      }${
        data["Havoc"]
          ? `\n|~|~&color(crimson){HC};|${data["Havoc_1"]}|${data["Havoc_2"]}|${
              data["Havoc"] === data["Rhythm"] ? "~" : data["Havoc"]
            }|`
          : ""
      }`;
    } else if (line.startsWith("|~BPM|")) {
      return `|~BPM|>|${data["BPM"]}|~Length|${data["曲の長さ"]}|`;
    } else if (line.startsWith("|~Version|")) {
      return `|~Version|>|>|>|${data["追加"]}|`;
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

    const excelData = readExcel("ryceam.xlsx");
    console.log("Loaded Excel data.");
    const targetChapters = ["Subject Spring / 重錦韶光"];
    for (const row of excelData) {
      const replacedName: string = EXCEPTIONS.reduce(
        (acc, exception) => acc.replace(exception.before, exception.after),
        row["曲名"]
      );
      if (IGNORED_PAGES.includes(replacedName)) {
        continue;
      }
      if (!targetChapters.includes(row["チャプター"])) {
        continue;
      }
      if (existingPages.includes(replacedName)) {
        const source = await wiki.getPage(replacedName);
        const updatedSource = updateWikiContent(source, row);
        await wiki.updatePage(replacedName, updatedSource);
      } else {
        const source = generateWikiContent(row);
        await wiki.updatePage(replacedName, source);
      }
      console.log(`Updated page: ${replacedName}`);
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
