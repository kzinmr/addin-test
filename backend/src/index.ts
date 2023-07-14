import express, { Request, Response } from "express";
import bodyParser from "body-parser";
import * as https from "https";
import * as devCerts from "office-addin-dev-certs";
import { Configuration, OpenAIApi } from "openai";

require('dotenv').config({ path: '.env.local' });

const app = express();
app.use(bodyParser.json());
const port = 9000;
const client = "https://localhost:3000";

// Handle Pre-flight Requests from the Browser
// See. https://developer.mozilla.org/en-US/docs/Web/HTTP/Methods/OPTIONS#preflighted_requests_in_cors
app.options("*", (req: Request, res: Response) => {
  res.header("Access-Control-Allow-Origin", client);
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept"
  );
  res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE");
  res.sendStatus(200);
});

const configuration = new Configuration({
    apiKey: process.env.OPENAI_API_KEY,
    organization: process.env.OPENAI_ORG_ID,
});
const openai = new OpenAIApi(configuration);
app.post("/ask", (req: Request, res: Response) => {
  res.header("Access-Control-Allow-Origin", client);
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept"
  );
  res.header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE");
  // check req.body has q key

  if (!configuration.apiKey) {
    res.status(500).json({
      error: {
        message: "OpenAI API key not configured.",
      },
    });
    return;
  }

  const q = req.body.q || "";
  if (q.trim().length === 0) {
    res.status(400).json({
      error: {
        message: "Please enter a valid query.",
      },
    });
    return;
  }

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
      // Consider adjusting the error handling logic for your use case
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
});

const options = async () => {
    return await devCerts.getHttpsServerOptions();
};
options().then((httpsOptions) => {
  https.createServer(httpsOptions, app).listen(port, () => {
    console.log("HTTPS Server running on port 9000");
  });
});
