import puppeteer from 'puppeteer';
import path from 'path';
import fs from 'fs';

(async () => {
  console.log('Avvio di Puppeteer...');
  const browser = await puppeteer.launch({ headless: false, protocolTimeout: 180000 });
  const page = await browser.newPage();

  // Imposta un viewport grande così da evitare menu collassati o layout mobile
  await page.setViewport({ width: 1400, height: 900 });

  // Gestione automatica di eventuali popup/alert
  page.on('dialog', async dialog => {
    console.log(`\n[Popup Rilevato] Messaggio: "${dialog.message()}"`);
    console.log(`[Popup] Lo script clicca "OK" automaticamente e procede...`);
    await dialog.accept();
  });
  
  const downloadPath = path.resolve('./temp_download');
  if (!fs.existsSync(downloadPath)) {
    fs.mkdirSync(downloadPath);
  }

  const client = await page.createCDPSession();
  await client.send('Page.setDownloadBehavior', {
    behavior: 'allow',
    downloadPath: downloadPath
  });

  // Funzione helper per attendere e rinominare il file scaricato con un timeout di sicurezza
  async function waitForDownloadAndMove(destPath, maxWaitSeconds = 60) {
    console.log(`In attesa che il file venga scaricato (timeout max: ${maxWaitSeconds}s)...`);
    let downloadedFile = null;
    let waited = 0;
    while (!downloadedFile && waited < maxWaitSeconds) {
      await new Promise(r => setTimeout(r, 1000));
      waited++;
      const files = fs.readdirSync(downloadPath);
      const finishedFile = files.find(f => f.length > 0 && !f.endsWith('.crdownload') && !f.endsWith('.tmp'));
      if (finishedFile) {
        downloadedFile = path.join(downloadPath, finishedFile);
      }
    }
    
    if (!downloadedFile) {
      console.log(`[AVVISO] Il file non è stato scaricato entro ${maxWaitSeconds} secondi. Procedo al prossimo step.`);
      return false;
    }
    
    if (fs.existsSync(destPath)) {
      try {
        fs.unlinkSync(destPath);
      } catch (err) {
        console.error(`Impossibile rimuovere il file esistente ${destPath}:`, err);
      }
    }
    try {
      fs.renameSync(downloadedFile, destPath);
      console.log(`Download completato con successo! Salvato in: ${destPath}\n`);
      return true;
    } catch (err) {
      console.error(`Errore nel rinominare il file in ${destPath}:`, err);
      return false;
    }
  }

  const exportSelector = '::-p-text(Esporta), ::-p-text(ESPORTA)';

  // Funzione helper per cliccare il pulsante Esporta e scaricare
  async function clickExportAndDownload(destPath, maxWaitSeconds = 60) {
    console.log('Attesa del pulsante "Esporta"...');
    try {
      const exportBtn = await page.waitForSelector(exportSelector, { timeout: 30000, visible: true });
      if (exportBtn) {
        console.log('Pulsante "Esporta" trovato! Clicco...');
        await page.evaluate(el => el.click(), exportBtn);
        // Aspetta per il file scaricato
        await waitForDownloadAndMove(destPath, maxWaitSeconds);
      }
    } catch (e) {
      console.log('[AVVISO] Nessun pulsante "Esporta" visibile in questa pagina o timeout scaduto. Procedo.');
    }
  }

  // Funzione helper per navigare in modo robusto tramite URL diretto
  async function navigateTo(url) {
    console.log(`Navigazione diretta verso: ${url}...`);
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    // Attesa di stabilità della pagina per AJAX/caricamenti interni
    await new Promise(r => setTimeout(r, 5000));
  }

  // Navigazione iniziale alla pagina delle chiamate (che provocherà il redirect al login se non loggati)
  console.log('Navigazione iniziale verso https://crm.askoltaora.it/office/calls...');
  await page.goto('https://crm.askoltaora.it/office/calls', { waitUntil: 'domcontentloaded' });
  await new Promise(r => setTimeout(r, 4000));

  const isLoginPage = await page.$('input[type="password"]');
  if (isLoginPage) {
    console.log('Rilevata pagina di login, inserimento credenziali in corso...');
    const usernameSelector = 'input[type="email"], input[type="text"], input[name="username"], input[name="email"]';
    await page.waitForSelector(usernameSelector);
    await page.type(usernameSelector, 'davide.coltrioli');
    await page.type('input[type="password"]', 'Edberg6!');
    console.log('Credenziali inserite. Invio del modulo e attesa caricamento...');
    await page.keyboard.press('Enter');
    
    // Attesa prolungata post-login vitale per le SPA
    await new Promise(r => setTimeout(r, 8000));
  }

  const desktopPath = 'C:\\Users\\dcolt\\Desktop';

  // --- PRIMO DOWNLOAD (Pagina Calls) ---
  console.log('\n--- Download 1: Calls (Chiamate) ---');
  await navigateTo('https://crm.askoltaora.it/office/calls');
  // Alziamo il timeout a 180 secondi perché la mole di chiamate è elevatissima
  await clickExportAndDownload(path.join(desktopPath, 'tt.xlsx'), 180);

  // --- SECONDO DOWNLOAD (Pagina Disponibilità) ---
  console.log('\n--- Download 2: Disponibilità ---');
  await navigateTo('https://crm.askoltaora.it/office/availabilities');
  await clickExportAndDownload(path.join(desktopPath, 'dd.xlsx'), 90);

  // --- TERZO DOWNLOAD (Pagina Appuntamenti) ---
  console.log('\n--- Download 3: Appuntamenti ---');
  await navigateTo('https://crm.askoltaora.it/office/appointments');
  await clickExportAndDownload(path.join(desktopPath, 'aa.xlsx'), 90);

  // --- QUARTO DOWNLOAD (Pagina Prove) ---
  console.log('\n--- Download 4: Prove ---');
  await navigateTo('https://crm.askoltaora.it/office/trials');
  await clickExportAndDownload(path.join(desktopPath, 'pp.xlsx'), 90);

  // --- QUINTO DOWNLOAD (Pagina Proforma) ---
  console.log('\n--- Download 5: Proforma ---');
  await navigateTo('https://crm.askoltaora.it/office/preinvoices');
  await clickExportAndDownload(path.join(desktopPath, 'ff.xlsx'), 90);

  // --- PULIZIA FINALE ---
  if (fs.existsSync(downloadPath)) {
    try {
      fs.rmSync(downloadPath, { recursive: true, force: true });
    } catch (e) {}
  }

  console.log('\nTutte le operazioni completate con successo! Chiusura del browser...');
  await browser.close();
})();
