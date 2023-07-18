/* global Word */
import React, { useState, useEffect } from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import Progress from "./Progress";

export type AppProps = {
  title: string;
  isOfficeInitialized: boolean;
};

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

  const [id, setId] = useState(null);
  const [eventSource, setEventSource] = useState(null);
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
    if (id && !eventSource) {
      const sseUrl = `https://localhost:9000/ask/sse/${id}`;
      const es = new EventSource(sseUrl);
      es.onerror = (error) => {
        console.error(error);
      };
      es.onmessage = (event) => {
        Word.run((context) => {
          context.document.body.insertText(event.data, Word.InsertLocation.end);
          return context.sync();
        }).catch((error) => {
          console.error(error);
        });
      };
      setEventSource(es); // effectの実行をトリガーするために状態を更新
    }

    return () => {
      if (eventSource) {
        eventSource.close(); // コンポーネント削除時にEventSourceも閉じる
      }
    };
  }, [id, eventSource]); // 状態が変更されたときにeffectを実行する

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
