document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('fileInput');
  const tabela = document.getElementById('tabelaExtrato');
  const thead = tabela.querySelector('thead');
  const tbody = tabela.querySelector('tbody');
  const resumo = document.getElementById('resumoTotais');
  const btnExport = document.getElementById('btnExport');

  // Esconde os elementos inicialmente
  resumo.style.display = 'none';
  btnExport.style.display = 'none';

  fileInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async function () {
      const typedArray = new Uint8Array(reader.result);
      const pdf = await pdfjsLib.getDocument({ data: typedArray }).promise;
      let todasLinhas = [];

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();

        let linhaAtual = '';
        let ultimaY = null;
        const toleranciaY = 4;

        content.items.forEach(item => {
          const y = Math.round(item.transform[5]);
          const texto = item.str.trim();

          if (ultimaY === null || Math.abs(y - ultimaY) <= toleranciaY) {
            linhaAtual += texto + ' ';
          } else {
            todasLinhas.push(linhaAtual.trim());
            linhaAtual = texto + ' ';
          }

          ultimaY = y;
        });

        if (linhaAtual.trim()) {
          todasLinhas.push(linhaAtual.trim());
        }
      }

      // Cabeçalho fixo com nova coluna D/C
      thead.innerHTML = '';
      const headerRow = document.createElement('tr');
      ['Data', 'Histórico', 'Valor', 'D/C'].forEach(coluna => {
        const th = document.createElement('th');
        th.textContent = coluna;
        headerRow.appendChild(th);
      });
      thead.appendChild(headerRow);

      // Extrair transações
      tbody.innerHTML = '';
      resumo.innerHTML = '';
      let ultimaData = '';
      const transacoes = [];

      todasLinhas.forEach(linha => {
        const dataMatch = linha.match(/\b\d{2}\/\d{2}\/\d{4}\b|\b\d{2}\/\d{2}\/\d{3}\b/);
        const valorMatch = linha.match(/-?\d{1,3}(?:\.\d{3})*,\d{2}(?: [CD-])?/);

        if (dataMatch) ultimaData = dataMatch[0];

        if (valorMatch) {
          const valorCompleto = valorMatch[0].trim();
          let tipo = 'C';

          if (
            valorCompleto.includes(' D') ||
            valorCompleto.includes(' -') ||
            valorCompleto.startsWith('-')
          ) {
            tipo = 'D';
          }

          const valorLimpo = valorCompleto
            .replace(/[^\d,]/g, '') // remove letras e sinais
            .replace(/\./g, '');   // remove pontos

          const valor = valorLimpo;
          const data = ultimaData || '';
          const inicio = data ? linha.indexOf(data) + data.length : 0;
          const fim = linha.lastIndexOf(valorCompleto);
          const historico = linha.substring(inicio, fim).trim();

          transacoes.push({ data, historico, valor, tipo });
        }
      });

      // Renderizar tabela na ordem original (sem ordenar por data)
      let totalC = 0;
      let totalD = 0;

      transacoes.forEach(({ data, historico, valor, tipo }) => {
        const valorNumerico = parseFloat(valor.replace(',', '.'));
        if (tipo === 'C') totalC += valorNumerico;
        if (tipo === 'D') totalD += valorNumerico;

        const tr = document.createElement('tr');
        [data, historico, valor, tipo].forEach(texto => {
          const td = document.createElement('td');
          td.textContent = texto;
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });

      // Exibir totalização fora da tabela
      const formatado = valor => valor.toLocaleString('pt-BR', { minimumFractionDigits: 2 });

      const totalCredito = document.createElement('p');
      totalCredito.innerHTML = `<strong>Total Crédito (C):</strong> R$ ${formatado(totalC)}`;

      const totalDebito = document.createElement('p');
      totalDebito.innerHTML = `<strong>Total Débito (D):</strong> R$ ${formatado(totalD)}`;

      const diferenca = document.createElement('p');
      diferenca.innerHTML = `<strong>Diferença (Débito - Crédito):</strong> R$ ${formatado(totalD - totalC)}`;

      resumo.appendChild(totalCredito);
      resumo.appendChild(totalDebito);
      resumo.appendChild(diferenca);

      // Mostrar os elementos
      resumo.style.display = 'block';
      btnExport.style.display = 'inline-block';
    };

    reader.readAsArrayBuffer(file);
  });

  // Botão para exportar para Excel
  document.getElementById('btnExport').addEventListener('click', () => {
    const tabela = document.getElementById('tabelaExtrato');
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(tabela);
    XLSX.utils.book_append_sheet(wb, ws, 'Extrato');
    XLSX.writeFile(wb, 'extrato_bancario.xlsx');
  });
});