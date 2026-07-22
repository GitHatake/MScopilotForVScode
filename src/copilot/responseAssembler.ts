/**
 * ストリーム本文の組み立てと、既知のサービス失敗フレーズの確実な検出を担う。
 * vscode/ネットワークに依存しない純ロジック(オフライン検証可能)。
 */

/**
 * サーバが「本文」として返す既知の一時的失敗メッセージ(実応答ではない)。
 * 誤ったスコープのトークンや合成リクエストが実テナントに合わず、モデル起動が拒否された
 * 場合に観測される。応答として採用せずエラー扱いにし、上位のトークン再取得リトライへ回す。
 */
export const SERVICE_FAILURE_TEXTS = ["language model unavailable"];
const MAX_FAILURE_LEN = Math.max(...SERVICE_FAILURE_TEXTS.map((s) => s.length));

/**
 * ストリーム本文を組み立てつつ、既知の失敗フレーズを確実に検出する。
 *
 * 失敗フレーズはスナップショットとして一括で届くとは限らず、writeAtCursor 差分で
 * 「Language model」→「 unavailable」のように分割して届くことがある。従来の
 * 「先頭が失敗フレーズか」判定だと最初の断片を本文として出してしまい、失敗が素の応答として
 * 表示され、リトライも発火しなかった(feedback_4 の症状)。
 *
 * 対策として、まだ何も出力していない間は「失敗フレーズの前方一致かつ十分短い」テキストを
 * バッファに留め、フレーズと異なる文字が来た時点で一括フラッシュする。最長フレーズ(~26文字)
 * を超える遅延は発生しない。ストリーム終了時に全文が失敗フレーズなら error 化する。
 */
export class ResponseAssembler {
  /** 受信済みの全文(保留分を含む)。 */
  private full = "";
  /** onDelta で実際に出力済みの文字数。 */
  private emittedLen = 0;
  /** 失敗フレーズと判定された場合の本文(あれば応答は無効)。 */
  failure?: string;

  constructor(private readonly onDelta: (delta: string) => void) {}

  /** 実際に本文を 1 文字でも出力したか(リトライ要否・close エラー判定に使う)。 */
  get produced(): boolean {
    return this.emittedLen > 0;
  }

  /** 応答本文。失敗フレーズなら空(実応答ではないため保存・表示しない)。 */
  get text(): string {
    return this.failure ? "" : this.full;
  }

  /** writeAtCursor 差分を末尾へ追記する。 */
  appendDelta(delta: string): void {
    if (!delta) {
      return;
    }
    this.full += delta;
    this.flush(false);
  }

  /** messages[].text の累積スナップショットを反映する。 */
  setSnapshot(snapshot: string): void {
    // 通常は前回全文の接頭辞が伸びる。稀に置換される場合は末尾へ続けて整合を保つ。
    this.full = snapshot.startsWith(this.full) ? snapshot : this.full + snapshot;
    this.flush(false);
  }

  /** ストリーム終了。保留を確定(失敗フレーズなら error 化、そうでなければ出力)する。 */
  end(): void {
    this.flush(true);
  }

  private flush(complete: boolean): void {
    if (this.failure) {
      return;
    }
    // まだ何も出力していない間だけ失敗フレーズ判定を行う(本物の応答を誤判定しないため)。
    if (this.emittedLen === 0) {
      const t = this.full.trim().toLowerCase();
      if (SERVICE_FAILURE_TEXTS.some((p) => p === t)) {
        this.failure = this.full.trim();
        return;
      }
      // 失敗フレーズの前方一致で、まだ短い間は確定を保留(分割到着に備える)。
      if (
        !complete &&
        this.full.length <= MAX_FAILURE_LEN &&
        t.length > 0 &&
        SERVICE_FAILURE_TEXTS.some((p) => p.startsWith(t))
      ) {
        return;
      }
    }
    // 保留を解除して未出力分をまとめて出す(full は常に増加し emittedLen 分は接頭辞)。
    const delta = this.full.slice(this.emittedLen);
    if (delta) {
      this.emittedLen = this.full.length;
      this.onDelta(delta);
    }
  }
}
