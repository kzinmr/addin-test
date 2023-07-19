
// @ts-nocheck  // for openai-node v3.0.x
import express, { Request, Response } from "express";
import bodyParser from "body-parser";
import * as https from "https";
import * as devCerts from "office-addin-dev-certs";
import { Configuration, OpenAIApi } from "openai";
import { v4 as uuidv4 } from 'uuid';
require('dotenv').config({ path: '.env.local' });

type ClientData = { questions: string[] }
const clientIds: Map<string, ClientData> = new Map();

// Set up OpenAI API Client
const configuration = new Configuration({
  apiKey: process.env.OPENAI_API_KEY,
  // organization: process.env.OPENAI_ORG_ID,
});
const openai = new OpenAIApi(configuration);

// Set up Express app
const app = express();
app.use(bodyParser.json());
app.use(setCORSHeaders);

function setCORSHeaders(req: Request, res: Response, next: NextFunction) {
  const client = "https://localhost:3000";
  res.header("Access-Control-Allow-Origin", client);
  res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
  res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE");
  next();
}

function checkOpenAIApiKey(req: Request, res: Response, next: NextFunction) {
  if (!configuration.apiKey) {
    res.status(500).json({
      error: { message: "OpenAI API key not configured." },
    });
    return;
  }
  next();
}

function checkRequestQuery(req: Request, res: Response, next: NextFunction) {
  const q = req.body.q || "";
  if (q.trim().length === 0) {
    res.status(400).json({
      error: { message: "Please enter a valid query." },
    });
    return;
  }
  next();
}

// Set up routes
app.options("*", (_, res: Response) => res.sendStatus(200));
// CORSリクエストの際にブラウザから送出されるPre-flight Requestsを許可するためのAPI。
// See. https://developer.mozilla.org/en-US/docs/Web/HTTP/Methods/OPTIONS#preflighted_requests_in_cors
app.post("/ask", checkOpenAIApiKey, checkRequestQuery, askHandler);
app.post("/ask/prepare", checkOpenAIApiKey, checkRequestQuery, askPrepareHandler);
app.get("/ask/sse/:id", askSSEHandler);

/**
 * ChatGPTからのレスポンスを通常のREST APIで返すAPI。
 * 結果のレスポンスはJSONで一気に返すため１０秒以上程度遅延が起こる。
 * @param onLine A function that will be called on each new EventSource line.
 * @returns ChatGPTの結果をJSONで返す。
 */
function askHandler(req: Request, res: Response) {
  const q = req.body.q || "";
  openai
    .createChatCompletion({
      model: "gpt-3.5-turbo-0613",
      messages: [
        {
          role: "system",
          content:
            "You are a helpful legal assistant who excels at drafting and reviewing contracts.",
        },
        { role: "user", content: q },
      ],
    })
    .then((completion) => {
      const content = completion.data.choices[0].message?.content;
      res.status(200).json({ result: content });
    })
    .catch((err) => {
      if (err) {
        console.error(err.response.status, err.response.data);
        res.status(err.response.status).json(err.response.data);
      } else {
        console.error(`Error with OpenAI API request: ${err.message}`);
        res.status(500).json({
          error: {
            message: "An error occurred during your request.",
          },
        });
      }
    });
};

/** 
 * ChatGPTからのレスポンスをServer-Sent Events(SSE)で返すための前準備のAPI。
 * SSEはリクエストBodyを受け取らずGETのみしか対応しないため、事前にパラメタ事前登録が必要。
 * リクエストに対してUUIDをキーに質問を保存し、SSEリクエスト時に取り出す。
 * @param JSONボディで質問を受け取る(key: q)。
 * @returns 発行したUUIDを返す。
 */
function askPrepareHandler(req: Request, res: Response) {
  const q = req.body.q || "";
  const id = uuidv4();
  clientIds.set(id, { questions: [q] });
  res.status(200).json({ id });
};

/** 
 * ChatGPTからのレスポンスをServer-Sent Events(SSE)で返すAPI。
 * 事前登録したChatGPTへのリクエストボディをPOSTし、返ってくるストリームをSSEで少しずつ返す。
 * @param id 事前リクエストで保存したUUID、ChatGPTへのリクエストを引き出すためのキー。
 * @returns ChatGPTの結果をSSEで返す。
 */
function askSSEHandler(req: Request, res: Response) {
  // socketのタイムアウトを無効化, todo: 適切な値を設定
  req.socket.setTimeout(10000);
  // レスポンスを Server-Sent Events として設定
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const sendEvents = (event: string, id: number, msg: string) => {
    const data = JSON.stringify({result: msg});
    // console.log(`event: ${event}\nid: ${id}\ndata: ${data}\n\n`);
    res.write(`event: ${event}\n`)
    res.write(`id: ${id}\n`);
    res.write(`data: ${data}\n\n`);
  };

  let buffer: string = '';
  let status: string | null = null;
  const intervalId = setInterval(async () => {
    const clientData = clientIds.get(req.params.id);
    if (typeof clientData !== "undefined" && clientData.questions.length > 0) {
      const question = clientData.questions.pop() || "";

      // todo: v4.0.0がリリースされたらStreamにSDK側が対応される見込み.
      // https://github.com/openai/openai-node/issues/18
      const completion = await openai.createChatCompletion({
        model: "gpt-3.5-turbo-0613",
        messages: [
          {
            role: "system",
            content:
              "You are a helpful legal assistant who excels at drafting and reviewing contracts.",
          },
          { role: "user", content: question },
        ],
        stream: true,
      }, { responseType: 'stream' })
      try {
        completion.data.on('data', (data) => {
          // console.log(data.toString());
          // 受け取ったイベントの data が完全なJSONオブジェクトを含まない場合に対処
          buffer += data.toString();
          const messages = buffer.split(/\ndata: /).filter(line => line.trim() !== '');
          // If the buffer does not end with a newline, then the last line is a partial line, so keep it in the buffer
          if (!buffer.endsWith('\n')) {
            buffer = messages.pop() || '';
          } else {
            buffer = '';
          }

          for (const message of messages) {
            const m: string = message.trim().replace(/^data: /, '');
            if (m === '[DONE]') {
              sendEvents('done', 0, '');
              if (status === 'done' && clientData.questions.length === 0) {
                console.log('close connection');
                res.end();
              }
              return;
            }
            try {
              const parsed = JSON.parse(m);
              const choice = parsed.choices[0];
              if (choice.finish_reason !== 'stop') {
                // 受信したイベントをクライアントに送信
                sendEvents('message', parsed.id, choice.delta?.content);
              }
            } catch(error) {
              console.error('Could not JSON parse stream message', m, error);
              console.error('Original line:', m, error);
            }
          }
        });
      } catch (err) {
        if (err.response?.status) {
          console.error(err.response.status, err.message);
          err.response.data.on('data', (data) => {
            res.status(err.response.status).json(data.toString());
          });
        } else {
          console.error('An error occurred during OpenAI request', err);
          res.status(500).json({
            error: {
              message: "An error occurred during your request.",
            },
          });          
        }
      };
    }
  }, 2000);

  req.on('close', () => {
    clearInterval(intervalId);
    clientIds.delete(req.params.id);
  });

};

// Office Add-inが必要とするHTTPSサーバーをlocalに建てるための措置
const options = async () => {
    return await devCerts.getHttpsServerOptions();
};
options().then((httpsOptions) => {
  https.createServer(httpsOptions, app).listen(9000, () => {
    console.log("HTTPS Server running on port 9000");
  });
});
