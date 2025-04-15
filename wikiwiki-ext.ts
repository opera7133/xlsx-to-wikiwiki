import axios from "axios";
import * as cheerio from "cheerio";

export class WikiWiki {
  private wikiName: string;
  private operationCount: number;
  private readonly maxOperationsPerMinute: number;

  constructor(wikiName: string, maxOperationsPerMinute: number = 120) {
    this.wikiName = wikiName;
    this.operationCount = 0;
    this.maxOperationsPerMinute = maxOperationsPerMinute;
  }

  private async checkRateLimit(): Promise<void> {
    if (this.operationCount >= this.maxOperationsPerMinute) {
      console.log("Rate limit reached. Waiting for 1 minute...");
      await new Promise((resolve) => setTimeout(resolve, 60000));
      this.operationCount = 0;
    }
  }

  async getPages(): Promise<string[]> {
    await this.checkRateLimit();
    const response = await axios.get(
      `https://wikiwiki.jp/${this.wikiName}/::cmd/list`
    );
    const $ = cheerio.load(response.data);
    this.operationCount++;
    return $("#content li a.rel-wiki-page")
      .map((_, element) => $(element).text())
      .get();
  }

  async getPage(pageName: string): Promise<string> {
    await this.checkRateLimit();
    const response = await axios.get(
      `https://wikiwiki.jp/${
        this.wikiName
      }/::cmd/source?page=${pageName.replaceAll(" ", "+")}`
    );
    const $ = cheerio.load(response.data);
    this.operationCount++;
    return $("pre#source").text();
  }
}
