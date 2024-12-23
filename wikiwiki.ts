import axios from "axios";

export class WikiWiki {
  private wikiName: string;
  private password: string;
  private token: string | null;
  private operationCount: number;
  private readonly maxOperationsPerMinute: number;

  constructor(
    wikiName: string,
    password: string,
    maxOperationsPerMinute: number = 120
  ) {
    this.wikiName = wikiName;
    this.password = password;
    this.token = null;
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

  async authenticate(): Promise<void> {
    await this.checkRateLimit();
    const response = await axios.post(
      `https://api.wikiwiki.jp/${this.wikiName}/auth`,
      {
        password: this.password,
      }
    );
    this.operationCount++;

    if (response.data.status === "ok") {
      this.token = response.data.token;
    } else {
      throw new Error("Failed to authenticate");
    }
  }

  async getPages(): Promise<string[]> {
    if (!this.token) {
      throw new Error("Not authenticated");
    }

    await this.checkRateLimit();
    const response = await axios.get(
      `https://api.wikiwiki.jp/${this.wikiName}/pages`,
      {
        headers: { Authorization: `Bearer ${this.token}` },
      }
    );
    this.operationCount++;
    return response.data.pages.map((page: any) => page.name);
  }

  async getPage(pageName: string): Promise<string> {
    if (!this.token) {
      throw new Error("Not authenticated");
    }

    await this.checkRateLimit();
    const response = await axios.get(
      `https://api.wikiwiki.jp/${this.wikiName}/page/${encodeURIComponent(
        pageName
      )}`,
      {
        headers: { Authorization: `Bearer ${this.token}` },
      }
    );
    this.operationCount++;
    return response.data.source;
  }

  async updatePage(pageName: string, source: string): Promise<void> {
    if (!this.token) {
      throw new Error("Not authenticated");
    }

    await this.checkRateLimit();
    const response = await axios.put(
      `https://api.wikiwiki.jp/${this.wikiName}/page/${encodeURIComponent(
        pageName
      )}`,
      { source },
      {
        headers: {
          Authorization: `Bearer ${this.token}`,
          "Content-Type": "application/json",
        },
      }
    );
    this.operationCount++;

    if (response.data.status !== "ok") {
      throw new Error(`Failed to update page: ${pageName}`);
    }
  }

  isAuthenticated(): boolean {
    return this.token !== null;
  }

  getCurrentOperationCount(): number {
    return this.operationCount;
  }
}
