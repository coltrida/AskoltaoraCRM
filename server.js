import express from 'express';
import cors from 'cors';
import { spawn } from 'child_process';
import path from 'path';

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

let isRunning = false;

app.post('/api/run-automation', (req, res) => {
  if (isRunning) {
    return res.status(400).json({ success: false, message: 'Lo script è già in esecuzione.' });
  }

  isRunning = true;
  console.log('Avvio script di automazione richiesto dal frontend...');

  // Esegue lo script usando node
  const scriptPath = path.resolve('./download_excel.js');
  const process = spawn('node', [scriptPath], {
    stdio: 'inherit' // Invia l'output direttamente alla console del server
  });

  process.on('close', (code) => {
    isRunning = false;
    console.log(`Script di automazione terminato con codice ${code}`);
  });

  process.on('error', (err) => {
    isRunning = false;
    console.error(`Errore nell'avvio dello script: ${err}`);
  });

  // Rispondiamo subito al frontend, il processo girerà in background
  res.json({ success: true, message: 'Script avviato in background.' });
});

app.get('/api/automation-status', (req, res) => {
  res.json({ isRunning });
});

app.listen(PORT, () => {
  console.log(`\n--- Server Automazione Iniziato ---`);
  console.log(`In ascolto sulla porta ${PORT}`);
  console.log(`Endpoint: POST http://localhost:${PORT}/api/run-automation\n`);
});
