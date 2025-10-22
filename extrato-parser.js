document.addEventListener('DOMContentLoaded', () => {
  // ✅ Corrigir uso do worker do PDF.js
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

  const fileInput = document.getElementById('fileInput');
  const tabela = document.getElementById('tabelaExtrato');
  const thead = tabela.querySelector('thead');
  const tbody = tabela.querySelector('tbody');
  const resumo = document.getElementById('resumoTotais');
  const btnExport = document.getElementById('btnExport');
  const btnFormatarHistorico = document.getElementById('btnFormatarHistorico');
  const btnRemoverLinhas = document.getElementById('btnRemoverLinhas');
  const btnRemoverTextoHistorico = document.getElementById('btnRemoverTextoHistorico');
  const removerContainer = document.getElementById('removerContainer');

  // Ocultar botões inicialmente
  [resumo, btnExport, btnFormatarHistorico, btnRemoverLinhas, btnRemoverTextoHistorico, removerContainer].forEach(el => {
    el.style.display = 'none';
  });

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

      // Cabeçalho
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
      let transacaoAtual = null;

      todasLinhas.forEach(linha => {
        const dataMatch = linha.match(/\b\d{2}\/\d{2}\/\d{4}\b/);
        const valorMatch = linha.match(/-?\d{1,3}(?:\.\d{3})*,\d{2}(?: [CD-])?/);

        if (dataMatch) ultimaData = dataMatch[0];

        if (valorMatch) {
          const valorCompleto = valorMatch[0].trim();
          let tipo = valorCompleto.includes(' D') || valorCompleto.includes(' -') || valorCompleto.startsWith('-') ? 'D' : 'C';

          const valor = valorCompleto.replace(/[^\d,]/g, '').replace(/\./g, '');
          const data = ultimaData || '';
          const inicio = data ? linha.indexOf(data) + data.length : 0;
          const fim = linha.lastIndexOf(valorCompleto);
          const historico = linha.substring(inicio, fim).trim();

          transacaoAtual = { data, historico, valor, tipo };
          transacoes.push(transacaoAtual);
        } else if (transacaoAtual) {
          transacaoAtual.historico += ' ' + linha.trim();
        }
      });

      // Renderizar tabela
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

      const formatado = valor => valor.toLocaleString('pt-BR', { minimumFractionDigits: 2 });

      resumo.innerHTML = `
        <p><strong>Total Crédito (C):</strong> R$ ${formatado(totalC)}</p>
        <p><strong>Total Débito (D):</strong> R$ ${formatado(totalD)}</p>
        <p><strong>Diferença (Débito - Crédito):</strong> R$ ${formatado(totalD - totalC)}</p>
      `;

      // Mostrar botões
      [resumo, btnExport, btnFormatarHistorico, btnRemoverLinhas, btnRemoverTextoHistorico, removerContainer].forEach(el => {
        el.style.display = 'block';
      });
    };

    reader.readAsArrayBuffer(file);
  });

  // Exportar para Excel
  btnExport.addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(tabela);
    XLSX.utils.book_append_sheet(wb, ws, 'Extrato');
    XLSX.writeFile(wb, 'extrato_bancario.xlsx');
  });

  // Formatar Histórico
  btnFormatarHistorico.addEventListener('click', () => {
    const linhas = tabela.querySelectorAll('tbody tr');
    const headers = tabela.querySelectorAll('thead th');
    let historicoIndex = Array.from(headers).findIndex(th => th.textContent.trim().toLowerCase() === 'histórico');
    if (historicoIndex === -1) return alert('Coluna "Histórico" não encontrada.');

    const removerAcentos = texto => texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

    linhas.forEach(linha => {
      const celulas = linha.querySelectorAll('td');
      const historicoCelula = celulas[historicoIndex];
      if (historicoCelula) {
        let texto = historicoCelula.textContent;
        texto = removerAcentos(texto).replace(/[0-9.]/g, '').toUpperCase();
        historicoCelula.textContent = texto.trim();
      }
    });
  });

  // Remover linhas com termos
  btnRemoverLinhas.addEventListener('click', () => {
    const textoFiltro = document.getElementById('filtroTexto').value.trim().toLowerCase();
    if (!textoFiltro) return;

    const termos = textoFiltro.split(',').map(t => t.trim()).filter(t => t.length > 0);
    const linhas = tabela.querySelectorAll('tbody tr');
    const headers = tabela.querySelectorAll('thead th');
    let historicoIndex = Array.from(headers).findIndex(th => th.textContent.trim().toLowerCase() === 'histórico');
    if (historicoIndex === -1) return alert('Coluna "Histórico" não encontrada.');

    linhas.forEach(linha => {
      const celulas = linha.querySelectorAll('td');
      const historicoCelula = celulas[historicoIndex];
      if (historicoCelula) {
        const texto = historicoCelula.textContent.toLowerCase();
        const deveRemover = termos.some(termo => texto.includes(termo));
        if (deveRemover) linha.remove();
      }
    });
  });

  // Remover parte do histórico
  btnRemoverTextoHistorico.addEventListener('click', () => {
    const textoFiltro = document.getElementById('filtroTexto').value.trim().toLowerCase();
    if (!textoFiltro) return;

    const termos = textoFiltro.split(',').map(t => t.trim()).filter(t => t.length > 0);
    const linhas = tabela.querySelectorAll('tbody tr');
    const headers = tabela.querySelectorAll('thead th');
    let historicoIndex = Array.from(headers).findIndex(th => th.textContent.trim().toLowerCase() === 'histórico');
    if (historicoIndex === -1) return alert('Coluna "Histórico" não encontrada.');

        linhas.forEach(linha => {
      const celulas = linha.querySelectorAll('td');
      const historicoCelula = celulas[historicoIndex];
      if (historicoCelula) {
        let texto = historicoCelula.textContent;
        termos.forEach(termo => {
          const regex = new RegExp(termo, 'gi');
          texto = texto.replace(regex, '');
        });
        historicoCelula.textContent = texto.trim();
      }
    });
  });
});