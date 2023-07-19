/* global Word */
import React, { useState, useEffect } from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import Progress from "./Progress";

export type AppProps = {
  title: string;
  isOfficeInitialized: boolean;
};

const wordRun = async (text: string) => {
  const lines = text.split("\n");
  Word.run(async (context) => {
    // バッチ処理由来の不自然な改行を避けるために、最初の行だけinsertTextで挿入する
    context.document.body.insertText(lines[0], Word.InsertLocation.end);
    for (const line of lines.slice(1)) {
      context.document.body.insertParagraph(line, Word.InsertLocation.end);
    }
    return context.sync();
  })
}

const App: React.FC<AppProps> = (props) => {
  if (!props.isOfficeInitialized) {
    return (
      <Progress
        title={props.title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  let eventSource = null;
  const [id, setId] = useState(null);
  const handleSubmit = (event) => {
    // フォームのデフォルトの送信動作をキャンセル
    event.preventDefault();

    // SSEはリクエストBodyを受け取らないので事前に登録する
    Word.run(async (context) => {
      const input = context.document.getSelection().paragraphs.getFirst();
      input.load("text");
      await context.sync();

      const response = await fetch("https://localhost:9000/ask/prepare", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ q: input.text }),
      });
      const responseData = await response.json();
      setId(responseData.id); // effectの実行をトリガーするために状態を更新
    });
  };

  useEffect(() => {
    let batchedMessages = [];
    const batchSize = 100;

    if (id) {
      let isProcessing = false;
      const maxRetries = 5;
      let currentRetries = 0;
      const sseUrl = `https://localhost:9000/ask/sse/${id}`;
      eventSource = new EventSource(sseUrl);
      eventSource.addEventListener("done", () => {
        eventSource.close();
        if (batchedMessages.length >= 0) {
          isProcessing = true;
          // flush batched messages
          const text = batchedMessages.join('');
          batchedMessages = [];
          try {
            wordRun(text);
          } catch (error) {
            console.error(error);
          } finally {
            isProcessing = false;
          }
        }
      });
      eventSource.onerror = (error) => {
        console.error(error);
        if (eventSource.readyState == EventSource.CLOSED) {
          eventSource.close();
        } else if (error.target.readyState === EventSource.CONNECTING) {
          if (currentRetries === maxRetries) {
            console.error("Max retries reached!");
            eventSource.close();
          } else {
            console.log("Connection error - retrying... (" + currentRetries + "/" + maxRetries + ")");
            currentRetries++;
          }
        } else {
          console.error("An unexpected error occurred: ", error);
        }
      };
      eventSource.onmessage = async (event) => {

        while(isProcessing) {
          // 現在の処理が終わるまで待つ
          await new Promise(resolve => setTimeout(resolve, 1000));
        }

        const d = JSON.parse(event.data.trim());
        batchedMessages.push(d.result);
        if (batchedMessages.length >= batchSize) {
          isProcessing = true;
          // flush batched messages
          const text = batchedMessages.join('');
          batchedMessages = [];
          try {
            wordRun(text);
          } catch (error) {
            console.error(error);
          } finally {
            isProcessing = false;
          }
        }
      };
    }

    return () => {
      if (eventSource) {
        eventSource.close(); // コンポーネント削除時にEventSourceも閉じる
      }
    };
  }, [id]); // 状態が変更されたときにeffectを実行する

  return (
    <div className="ms-welcome">
      <Header logo={require("./../../../assets/logo-filled.png")} title={props.title} message="Ask Me" />

      <main className="ms-welcome__main">
        <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">Ask LLM</h2>

        <p className="ms-font-l">
          Select the body text, then click <b>Ask</b>.
        </p>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={handleSubmit}>
          Ask
        </DefaultButton>
      </main>
    </div>
  );
};

export default App;
