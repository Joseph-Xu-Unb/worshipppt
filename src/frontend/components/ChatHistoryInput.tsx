import { useId, useState } from "react";

export function ChatHistoryInput({ instructions, assistant }: { instructions: string; assistant: string }) {
  const id = useId();
  const [history, setHistory] = useState("");
  const [message, setMessage] = useState("");
  const [showFallback, setShowFallback] = useState(false);
  const combinedText = `${instructions}\n\n以下是教会微信群或 WhatsApp 群的聊天记录：\n${history.trim()}`;

  async function copy() {
    try {
      await navigator.clipboard.writeText(combinedText);
      setMessage(`已复制，请粘贴到 ${assistant} 并发送。`);
      setShowFallback(false);
    } catch {
      setMessage("请选中下方合并后的文字，手动复制。");
      setShowFallback(true);
    }
  }

  return (
    <div>
      <label htmlFor={id}><strong>在这里粘贴教会群聊天记录</strong></label>
      <textarea
        id={id}
        className="instructions-box"
        value={history}
        onChange={(event) => {
          setHistory(event.target.value);
          setMessage("");
        }}
        placeholder="请粘贴本次崇拜的筹备消息：聚会日期、经文出处、诗歌名称、完整歌词及后续更正……"
      />
      <button type="button" disabled={!history.trim()} onClick={() => void copy()}>
        复制指令和聊天记录
      </button>
      <p role="status" aria-live="polite">{message}</p>
      <details open={showFallback || undefined}>
        <summary>查看合并后的文字</summary>
        <textarea className="instructions-box" value={combinedText} readOnly aria-label="指令和聊天记录" />
      </details>
    </div>
  );
}
