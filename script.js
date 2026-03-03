const pdfjsLib = window['pdfjs-dist/build/pdf'];
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

document.getElementById('pdfInput').addEventListener('change', function() {
    const fileListContainer = document.getElementById('file-list');
    fileListContainer.innerHTML = ''; 
    
    if (this.files.length > 0) {
        for (let i = 0; i < this.files.length; i++) {
            const p = document.createElement('p');
            p.innerHTML = `📄 <strong>${this.files[i].name}</strong>`;
            fileListContainer.appendChild(p);
        }
    }
});

async function processFilesAndExport(format) {
    const files = document.getElementById('pdfInput').files;
    if (files.length === 0) return alert("Selecione pelo menos um PDF");

    const statusDiv = document.getElementById('status');
    let allPdfsData = []; 

    for (let f = 0; f < files.length; f++) {
        const file = files[f];
        statusDiv.innerHTML = `Iniciando arquivo ${f + 1} de ${files.length}: ${file.name}...`;

        const fileData = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = async function() {
                try {
                    const typedarray = new Uint8Array(this.result);
                    const pdf = await pdfjsLib.getDocument(typedarray).promise;
                    let fullData = [];
                    
                    let orfaNumber = file.name.match(/\d+[\/-]\d+/) ? file.name.match(/\d+[\/-]\d+/)[0] : file.name.replace('.pdf', '');

                    for (let i = 1; i <= pdf.numPages; i++) {
                        statusDiv.innerHTML = `Arquivo ${f + 1}/${files.length} | Processando página ${i} de ${pdf.numPages}...`;
                        const page = await pdf.getPage(i);
                        const textContent = await page.getTextContent();
                        
                        const items = textContent.items.filter(item => item.str.trim().length > 0);
                        let pageText = "";
                        let charToItemMap = [];

                        items.forEach((item) => {
                            let str = item.str.trim();
                            let startIdx = pageText.length;
                            pageText += str + " ";
                            
                            for (let c = startIdx; c < pageText.length; c++) {
                                charToItemMap.push({ x: item.transform[4] });
                            }
                        });

                        const expositorRegex = /((?:\d{1,2}\s+)?[A-Z0-9]{2,}(?:\s+[A-Z0-9\-]+)*\s+\d{3,4}\s*\([\d\.\-\s]+\))/g;
                        let expositores = [];
                        let match;
                        while ((match = expositorRegex.exec(pageText)) !== null) {
                            expositores.push({ 
                                name: match[0].trim(), 
                                start: match.index, 
                                end: match.index + match[0].length 
                            });
                        }

                        for (let k = 0; k < expositores.length; k++) {
                            const exp = expositores[k];
                            const prevExp = expositores[k - 1];
                            const nextExp = expositores[k + 1];

                            let sliceStart = prevExp ? Math.floor((prevExp.end + exp.start) / 2) : 0;
                            let sliceEnd = nextExp ? Math.floor((exp.end + nextExp.start) / 2) : pageText.length;

                            if (k === 0) sliceStart = Math.max(0, exp.start - 600); 

                            let rawSlice = pageText.substring(sliceStart, sliceEnd);

                            let dirFinal = "Vazio";
                            let esqFinal = "Vazio";
                            let resultadoUnificado = "";

                            const idxDir = rawSlice.indexOf("Lateral Direita");
                            const idxEsq = rawSlice.indexOf("Lateral Esquerda");
                            const idxDiv = rawSlice.indexOf("Div. Interna");

                            const isTableLayout = (idxDir !== -1 && idxEsq !== -1 && Math.abs(idxEsq - idxDir) < 150);

                            if (isTableLayout && idxDiv !== -1) {
                                let contentStart = idxDiv + "Div. Interna".length;
                                
                                let stops = [
                                    rawSlice.indexOf("Rodapé", contentStart),
                                    rawSlice.indexOf("Acabamento", contentStart),
                                    rawSlice.indexOf("Acb.", contentStart)
                                ].filter(x => x !== -1);
                                
                                let contentEnd = stops.length > 0 ? Math.min(...stops) : rawSlice.length;
                                
                                let rowContent = rawSlice.substring(contentStart, contentEnd).trim();
                                rowContent = rowContent.replace(/["]/g, ""); 

                                rowContent = cleanGarbage(rowContent);

                                let parsed = parseByKeywords(rowContent);
                                dirFinal = parsed.dir;
                                esqFinal = parsed.esq;

                                if (dirFinal !== "Vazio" && esqFinal === "Vazio" && !/^\s*,/.test(rowContent)) {
                                    let searchBase = dirFinal.replace("RETA ", "").split(' ')[0];
                                    if (searchBase.length < 3 && dirFinal.length >= 5) searchBase = dirFinal.substring(0, 5);
                                    
                                    let wordIdx = rawSlice.indexOf(searchBase, contentStart); 
                                    
                                    if (wordIdx !== -1) {
                                        let globalPos = sliceStart + wordIdx;
                                        let itemData = charToItemMap[globalPos];
                                        
                                        if (itemData) {
                                            let wordX = itemData.x;
                                            
                                            let dirMap = charToItemMap[sliceStart + idxDir];
                                            let esqMap = charToItemMap[sliceStart + idxEsq];
                                            
                                            if (dirMap && esqMap) {
                                                let dirX = dirMap.x;
                                                let esqX = esqMap.x;
                                                let midpoint = (dirX + esqX) / 2;
                                                
                                                if (dirX < esqX && wordX > midpoint) {
                                                    esqFinal = dirFinal;
                                                    dirFinal = "Vazio";
                                                }
                                            }
                                        }
                                    }
                                }

                            } else {
                                if (idxDir !== -1) {
                                    let start = idxDir + "Lateral Direita".length;
                                    let stops = [
                                        rawSlice.indexOf("Rodapé", start),
                                        rawSlice.indexOf("Div.", start),
                                        rawSlice.indexOf("Acabamento", start),
                                        rawSlice.indexOf("Informações", start),
                                        (idxEsq > start) ? idxEsq : -1 
                                    ].filter(x => x > start);
                                    
                                    let end = stops.length > 0 ? Math.min(...stops) : rawSlice.length;
                                    dirFinal = rawSlice.substring(start, end).trim();
                                }

                                if (idxEsq !== -1) {
                                    let start = idxEsq + "Lateral Esquerda".length;
                                    let stops = [
                                        rawSlice.indexOf("Rodapé", start),
                                        rawSlice.indexOf("Div.", start),
                                        rawSlice.indexOf("Acabamento", start),
                                        rawSlice.indexOf("Informações", start),
                                        (idxDir > start) ? idxDir : -1 
                                    ].filter(x => x > start);
                                    
                                    let end = stops.length > 0 ? Math.min(...stops) : rawSlice.length;
                                    esqFinal = rawSlice.substring(start, end).trim();
                                }
                            }

                            dirFinal = cleanText(dirFinal).replace(/(.*)\b\1\b/g, '$1').trim();
                            esqFinal = cleanText(esqFinal).replace(/(.*)\b\1\b/g, '$1').trim();

                            if (dirFinal.endsWith("RETA")) {
                                dirFinal = dirFinal.substring(0, dirFinal.length - 4).trim();
                                if (dirFinal === "") dirFinal = "Vazio";
                                
                                if (esqFinal !== "Vazio" && !esqFinal.startsWith("RETA")) {
                                    esqFinal = "RETA " + esqFinal;
                                }
                            }

                            resultadoUnificado = `${dirFinal} (Direita) | ${esqFinal} (Esquerda)`;

                            const matExtMatch = rawSlice.match(/[Aa]c[bh]\.\s*[Ee]xterno:\s*([A-Z0-9]+(?:\s+[A-Z0-9]+)*\s*(?:\([A-Z0-9]+\))?)/);
                            const matIntMatch = rawSlice.match(/[Aa]cb\.\s*TT\/PNL:\s*([A-Z0-9]+(?:\s+[A-Z0-9]+)*\s*(?:\([A-Z0-9]+\))?)/);

                            fullData.push({
                                "ORFA": orfaNumber,
                                "Expositor": exp.name,
                                "Laterais": resultadoUnificado,
                                "Material Externo": matExtMatch ? matExtMatch[1].trim() : "PRETO (P1902)",
                                "Material Interno": matIntMatch ? matIntMatch[1].trim() : "INOX"
                            });
                        }
                    }
                    resolve(fullData); 
                } catch (error) {
                    console.error("Erro no arquivo " + file.name, error);
                    resolve([]); 
                }
            };
            reader.readAsArrayBuffer(file);
        });

        allPdfsData = allPdfsData.concat(fileData);
    } 

    const uniqueData = allPdfsData.filter((v, i, a) => 
        a.findIndex(t => t.Expositor === v.Expositor && t["Laterais"] === v["Laterais"] && t["ORFA"] === v["ORFA"]) === i
    );

    if (uniqueData.length === 0) {
        statusDiv.innerHTML = "Aviso: Nenhum expositor foi encontrado nos PDFs selecionados.";
        alert("Nenhum expositor encontrado.");
    } else {
        statusDiv.innerHTML = `Processamento concluído! Baixando ${format.toUpperCase()} com dados de ${files.length} arquivo(s)...`;
        
        if (format === 'excel') {
            exportToExcel(uniqueData);
        } else if (format === 'pdf') {
            exportToPdf(uniqueData);
        }
        
        setTimeout(() => { statusDiv.innerHTML = ""; }, 4000);
    }
}

document.getElementById('processBtn').addEventListener('click', () => processFilesAndExport('excel'));
document.getElementById('processPdfBtn').addEventListener('click', () => processFilesAndExport('pdf'));

function cleanGarbage(text) {
    const stopWords = ["DIVISORIAS", "DIVISORIA", "PRATELEIRAS", "PRATELEIRA", "GANCHEIRA", "GANCHEIRAS", "SUPORTE", "ILUMINADAS", "VIDRO"];
    let cutIndex = text.length;
    stopWords.forEach(word => {
        const regex = new RegExp(`(^|\\s)${word}`, 'i');
        const match = text.match(regex);
        if (match && match.index < cutIndex) cutIndex = match.index;
    });
    return text.substring(0, cutIndex).trim();
}

function cleanText(text) {
    if (!text || text.length < 3 || text.includes("Lateral")) return "Vazio";
    let limpo = text.replace(/\s0\d(\s|$)/g, " ").replace(/\s\d{2}(\s|$)/g, " ");
    limpo = limpo.replace(/^[,.\s]+/, "").replace(/[,.\s]+$/, "");
    limpo = cleanGarbage(limpo);
    return limpo.length > 2 ? limpo : "Vazio";
}

function parseByKeywords(rawContent) {
    const keywords = ["PANORAMICA", "CEGA", "INTERMEDIARIA", "CHANFRADA", "FECHADA", "VAZIO", "Vazio"];
    let matches = [];
    keywords.forEach(key => {
        const regex = new RegExp(`(^|\\s)(${key})`, 'gi');
        let m;
        while ((m = regex.exec(rawContent)) !== null) {
            matches.push({ index: m.index + m[1].length, word: m[2].toUpperCase() });
        }
    });

    matches.sort((a, b) => a.index - b.index);
    matches = matches.filter((v, i, a) => i === 0 || v.index !== a[i-1].index); 

    for (let k = 0; k < matches.length - 1; k++) {
        const cur = matches[k];
        const nxt = matches[k+1];
        if ((cur.word === "CHANFRADA" && nxt.word === "CEGA") || 
            (cur.word === "INTERMEDIARIA" && nxt.word === "CHANFRADA")) {
            if (nxt.index - (cur.index + cur.word.length) < 10) {
                matches.splice(k+1, 1);
                k--;
            }
        }
    }

    let dir = "Vazio";
    let esq = "Vazio";

    if (matches.length >= 2) {
        let split = matches[1].index;
        dir = rawContent.substring(0, split).trim();
        esq = rawContent.substring(split).trim();
    } else if (matches.length === 1) {
        if (/^\s*,/.test(rawContent)) { 
            esq = rawContent.replace(/^\s*,/, "").trim();
        } else {
            dir = rawContent.trim();
        }
    } else {
        let parts = rawContent.split(",");
        dir = parts[0] ? parts[0].trim() : "Vazio";
        esq = (parts.length > 1) ? parts[1].trim() : "Vazio";
    }
    return { dir, esq };
}

function exportToExcel(data) {
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{wch: 15}, {wch: 40}, {wch: 90}, {wch: 20}, {wch: 15}];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Producao");
    XLSX.writeFile(wb, "Relatorio_Producao_Multiplo.xlsx");
}

// --- FUNÇÃO DO PDF TOTALMENTE REFEITA PARA FICAR PROFISSIONAL ---
function exportToPdf(data) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('landscape'); 

    // Cores NSF
    const brandBlue = [0, 86, 179];
    const brandDark = [51, 51, 51];
    const brandGray = [100, 100, 100];

    // --- CABEÇALHO DO PDF ---
    doc.setFillColor(...brandBlue);
    doc.rect(0, 0, doc.internal.pageSize.width, 25, 'F'); // Faixa azul no topo

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(22);
    doc.setFont("helvetica", "bold");
    doc.text("NSF", 14, 16);
    
    doc.setFontSize(14);
    doc.setFont("helvetica", "normal");
    doc.text("| Relatório de Extração de ORFA (LATERAIS)", 32, 16);

    // --- INFORMAÇÕES DO DOCUMENTO ---
    doc.setTextColor(...brandDark);
    doc.setFontSize(10);
    const dataAtual = new Date().toLocaleDateString('pt-BR');
    const horaAtual = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
    doc.text(`Data de Geração: ${dataAtual} às ${horaAtual}`, 14, 33);
    doc.text(`Total de Expositores Extraídos: ${data.length}`, 14, 38);

    // --- CONFIGURAÇÃO DA TABELA ---
    const columns = ["ORFA", "Expositor", "Laterais", "Material Externo", "Material Interno"];
    const rows = data.map(item => [
        item["ORFA"],
        item["Expositor"],
        item["Laterais"],
        item["Material Externo"],
        item["Material Interno"]
    ]);

    doc.autoTable({
        startY: 45,
        head: [columns],
        body: rows,
        theme: 'grid',
        styles: { 
            font: 'helvetica',
            fontSize: 8, 
            cellPadding: 4,
            textColor: brandDark,
            lineColor: [220, 220, 220],
            lineWidth: 0.1,
        },
        headStyles: { 
            fillColor: brandBlue,
            textColor: 255,
            fontStyle: 'bold',
            halign: 'center'
        },
        alternateRowStyles: {
            fillColor: [248, 249, 250] // Linhas zebradas bem suaves
        },
        columnStyles: {
            0: { cellWidth: 22, halign: 'center', fontStyle: 'bold' },
            1: { cellWidth: 70 },
            2: { cellWidth: 110 },
            3: { cellWidth: 35 },
            4: { cellWidth: 30 }
        },
        didDrawPage: function (data) {
            // --- RODAPÉ COM NÚMERO DE PÁGINAS ---
            const pageCount = doc.internal.getNumberOfPages();
            const pageSize = doc.internal.pageSize;
            const pageHeight = pageSize.height ? pageSize.height : pageSize.getHeight();
            
            doc.setFontSize(8);
            doc.setTextColor(...brandGray);
            
            // Texto à esquerda
            doc.text('Documento gerado automaticamente pelo Extrator NSF', 14, pageHeight - 10);
            
            // Paginação à direita
            doc.text(`Página ${data.pageNumber}`, pageSize.width - 25, pageHeight - 10);
        }
    });

    doc.save("Relatorio_Producao_NSF.pdf");
}