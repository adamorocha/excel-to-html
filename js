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
