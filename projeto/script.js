// Função para carregar a imagem QR Code e retornar uma Promise com o Data URL
function carregarImagem(idImg) {
    return new Promise((resolve, reject) => {
        const img = document.getElementById(idImg);
        if (!img) {
            reject(new Error(`Imagem com id ${idImg} não encontrada.`));
        }

        // Cria um canvas para desenhar a imagem e obter o Data URL
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');

        img.onload = () => {
            canvas.width = img.width;
            canvas.height = img.height;
            ctx.drawImage(img, 0, 0);
            const dataURL = canvas.toDataURL('image/jpeg'); // Pode ajustar para 'image/png' se preferir
            resolve(dataURL);
        };

        img.onerror = () => {
            reject(new Error('Erro ao carregar a imagem.'));
        };

        // Se a imagem já estiver carregada
        if (img.complete) {
            img.onload();
        }
    });
}

document.getElementById('btnGerar').addEventListener('click', function () {
    const input = document.getElementById('inputExcel');
    if (!input.files.length) {
        alert('Por favor, selecione um arquivo Excel.');
        return;
    }

    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = async function (e) {
        try {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheet = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheet];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);

            // Carregar a imagem QR Code
            const qrcodeDataURL = await carregarImagem('qrcode');

            // Criar um novo ZIP
            const zip = new JSZip();

            // Gerar PDFs para cada linha e adicioná-los ao ZIP
            for (let index = 0; index < jsonData.length; index++) {
                const row = jsonData[index];
                const pdfDoc = gerarPDF(row, index + 1, qrcodeDataURL);
                const pdfBlob = await pdfDoc.output('blob');
                const nomeArquivo = `bilhete_${row.nome || 'sem_nome'}.pdf`; // Usar dados.nome para o nome do arquivo
                zip.file(nomeArquivo, pdfBlob);
            }

            // Gerar o arquivo ZIP e iniciar o download
            const content = await zip.generateAsync({ type: 'blob' });
            saveAs(content, 'bilhetes.zip');
        } catch (error) {
            console.error(error);
            alert('Ocorreu um erro ao gerar os PDFs.');
        }
    };

    reader.onerror = function (ex) {
        console.log(ex);
    };

    reader.readAsBinaryString(file);
});

function gerarPDF(dados, indice, qrcodeDataURL) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // Definir estilos e posições
    let y = 20;
    doc.setFontSize(24);
    doc.setFont('Helvetica', 'bold');
    doc.text('BILHETE DE PASSAGEM', 105, y, { align: 'center' });
    y += 5;
    doc.setLineWidth(0.5);
    doc.line(20, y, 190, y);
    y += 10;

    // Dados do cliente
    doc.setFontSize(20);
    doc.setFont('Helvetica', 'normal');
    doc.text(`Nome: ${dados.nome || ''}`, 20, y);
    y += 7;
    doc.text(`Nº do Documento: ${dados.numero_documento || ''}`, 20, y);
    y += 7;
    doc.text(`Saída: ${dados.saida || ''}`, 20, y);
    y += 7;
    doc.text(`Destino: ${dados.destino || ''}`, 20, y);

    // Tabela com Data da Viagem, Horário, Agência, Poltrona
    y += 15;
    doc.text(`Data da Viagem: ${formatarData(dados.data_viagem)}`, 20, y);
    const horario = formatarHorario(dados.horario);
    doc.text(`Horário: ${horario}`, 80, y);
    doc.text(`Agente: ${dados.agente || ''}`, 120, y);
    doc.text(`Poltrona: ${dados.poltrona || ''}`, 170, y);

    // Tabela com Data de Emissão, Valor, Agente ou Agência
    y += 15;
    doc.text(`Data de Emissão: ${formatarData(dados.data_emissao)}`, 20, y);
    doc.text(`Valor: R$ ${dados.valor || ''}`, 80, y);
    doc.text(`Agência: Anny Viagens`, 120, y);

    // Texto adicional e Fone
    y += 15;
    doc.text(`FONE: ${dados.fone || ''}`, 20, y);
    y += 15;
    const texto = "Para adiar a passagem o passageiro deverá nos avisar com 02 dias de antecedência, o passageiro tem direito a uma bolsa e uma mala, acima disso será cobrado a parte.";
    const textoQuebrado = doc.splitTextToSize(texto, 170);
    doc.text(textoQuebrado, 20, y);

    // Mensagem final
    y += 15;
    doc.setFont('Helvetica', 'bold');
    doc.text('POR FAVOR CHEGAR 1 HORA ANTES DO EMBARQUE * BOA VIAGEM', 105, y, { align: 'center' });

    // Adicionar a imagem QR Code no final do PDF
    const imgWidth = 50; // Largura da imagem em mm
    const imgHeight = 50; // Altura da imagem em mm
    const imgX = (doc.internal.pageSize.getWidth() - imgWidth) / 2; // Centralizar horizontalmente
    const imgY = y + 10; // Espaçamento após a mensagem final

    doc.addImage(qrcodeDataURL, 'JPEG', imgX, imgY, imgWidth, imgHeight);

    // Retornar o documento PDF para ser adicionado ao ZIP
    return doc;
}

function formatarHorario(horarioExcel) {
    if (typeof horarioExcel === 'number') {
        const totalSeconds = Math.round(horarioExcel * 24 * 60 * 60);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }
    return horarioExcel || '';
}

function formatarData(dataExcel) {
    if (!dataExcel) return '';
    let dataObj;
    if (typeof dataExcel === 'number') {
        // Se for um número, converte do formato Excel para data
        dataObj = new Date((dataExcel - 25568) * 86400 * 1000);
    } else {
        // Se já for uma data, converte para objeto Date
        dataObj = new Date(dataExcel);
    }
    const dia = String(dataObj.getDate()).padStart(2, '0');
    const mes = String(dataObj.getMonth() + 1).padStart(2, '0');
    const ano = dataObj.getFullYear();
    return `${dia}/${mes}/${ano}`;
}
