const xlsx = require('xlsx');
const path = require('path');

const caData = new Map();
let isDataReady = false;

function loadData() {
  try {
    console.log('[DADOS] Iniciando carregamento do arquivo caepi.xlsm...');
    const filePath = path.resolve(__dirname, 'caepi.xlsm');
    
    const workbook = xlsx.readFile(filePath);
    
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    const records = xlsx.utils.sheet_to_json(worksheet);
    
    caData.clear();
    for (const record of records) {
      const caKey = record['CA'];
      if (caKey) {
        caData.set(String(caKey).trim(), record);
      }
    }
    
    isDataReady = true;
    console.log(`[DADOS] Base de dados carregada com sucesso. Total de ${caData.size} registros.`);

  } catch (err) {
    console.error('[DADOS] ERRO AO LER O ARQUIVO EXCEL (.xlsm):', err.message);
  }
}

function getCAInfo(caNumber) {
  if (!isDataReady) {
    return { error: 'A base de dados ainda está a ser carregada. Por favor, tente novamente em um minuto.' };
  }
  
  const caInfo = caData.get(String(caNumber).trim());
  if (caInfo) {
    return {
      'Nº do CA': caInfo['CA'],
      'Data de Validade': caInfo['Validade'],
      'Situação': caInfo['Situacao'],
      'Equipamento': caInfo['Equipamento'],
      'Fabricante': caInfo['Fabricante']
    };
  } else {
    return { error: `O CA "${caNumber}" não foi encontrado na base de dados.` };
  }
}

module.exports = { getCAInfo, loadData };