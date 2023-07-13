import React, { Component } from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import Progress from "./Progress";

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {}

async function ask() {
  return Word.run(async (context) => {
    const input = context.document.getSelection().paragraphs.getFirst();
    input.load("text");
    await context.sync();
    console.log(input.text);

    const endpoint = "https://localhost:9000/ask";
    try {
      const received: string = await submit(endpoint, input.text);
      const paragraph = context.document.body.insertText(received, Word.InsertLocation.end);
      await context.sync();
    } catch (error) {
      console.error(error);
    }
  });
}

async function submit(url = "", data = ""): Promise<string> {
  const response = await fetch(url, {
    method: "POST",
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({q: data}),
  });
  if (!response.ok) {
    throw new Error(`HTTP error, status = ${response.status}`);
  }
  console.log(response);
  return response.text();
}

export default class App extends Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {};
  }

  click = ask;

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Ask Me" />

        <main className="ms-welcome__main">
          <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">Ask LLM</h2>

          <p className="ms-font-l">
            Select the body text, then click <b>Ask</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Ask
          </DefaultButton>
        </main>
      </div>
    );
  }
}
