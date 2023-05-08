// Importando a biblioteca SheetJS
import * as XLSX from 'xlsx';

// Selecionando o input do tipo file
const input = document.querySelector('input[type=file]');

// Adicionando um evento de mudança para quando o arquivo for selecionado
input.addEventListener('change', () => {
  // Criando um novo objeto do tipo FileReader
  const reader = new FileReader();
  
  // Adicionando um evento de carregamento para quando o arquivo for lido
  reader.onload = (e) => {
    // Obtendo os dados do arquivo
    const data = e.target.result;
    
    // Convertendo o arquivo para um objeto do tipo workbook do SheetJS
    const workbook = XLSX.read(data, { type: 'binary' });
    
    // Selecionando a primeira planilha do workbook
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    
    // Convertendo a planilha para um objeto JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Fazendo algo com os dados, como exibi-los na tela
    console.log(jsonData);
  };
  
  // Lendo o arquivo como um arquivo binário
  reader.readAsBinaryString(input.files[0]);
});

// Importando a biblioteca SheetJS
import * as XLSX from 'xlsx';

// Obtendo o input do usuário (arquivo do Excel)
const input = document.querySelector('input[type="file"]');

// Escutando o evento de mudança no input do usuário
input.addEventListener('change', () => {
  // Lendo o arquivo do Excel selecionado
  const file = input.files[0];
  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, {type: 'array'});
    // Obtendo a primeira planilha
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    // Convertendo as linhas da planilha em imagens
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    rows.forEach((row) => {
      const img = document.createElement
