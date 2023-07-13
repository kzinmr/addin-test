import express, { Request, Response } from 'express';
import bodyParser from "body-parser";
import * as https from 'https';
import * as devCerts from 'office-addin-dev-certs';

const options = async () => {
    return await devCerts.getHttpsServerOptions();
};

const app = express();
const port = 9000;

app.use(bodyParser.json());
// Handle Pre-flight Requests from the Browser
app.options('*', (req, res) => {
    res.header('Access-Control-Allow-Origin', 'https://localhost:3000');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
    res.sendStatus(200);
});
app.post('/ask', (req: Request, res: Response) => {
    res.header('Access-Control-Allow-Origin', 'https://localhost:3000');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
    // check req.body has q key
    if (req.body.q !== undefined) {
        res.send({ "Echoed data": req.body.q });
    } else {
        res.status(400).send({ error: 'Request body missing the q parameter' });
    }
  });

options().then(httpsOptions => {
    https.createServer(httpsOptions, app).listen(port, () => {
        console.log('HTTPS Server running on port 9000');
    });
})
