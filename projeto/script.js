document.getElementById('btnGerar').addEventListener('click', function () {
    const input = document.getElementById('inputExcel');
    if (!input.files.length) {
        alert('Por favor, selecione um arquivo Excel.');
        return;
    }

    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'binary' }); // Adicionado
        const firstSheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheet];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        // Criar um novo ZIP
        const zip = new JSZip();

        // Gerar PDFs para cada linha e adicioná-los ao ZIP
        jsonData.forEach((row, index) => {
            const pdfDoc = gerarPDF(row, index + 1);
            zip.file(`bilhete_${index + 1}.pdf`, pdfDoc.output('blob'));
        });

        // Gerar o arquivo ZIP e iniciar o download
        zip.generateAsync({ type: 'blob' }).then(function (content) {
            saveAs(content, 'bilhetes.zip');
        });
    };

    reader.onerror = function (ex) {
        console.log(ex);
    };

    reader.readAsBinaryString(file);
});

function gerarPDF(dados, indice) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // Definir estilos e posições
    let y = 20;
    doc.setFontSize(16);
    doc.setFont('Helvetica', 'bold');
    doc.text('BILHETE DE PASSAGEM', 105, y, { align: 'center' });
    y += 5;
    doc.setLineWidth(0.5);
    doc.line(20, y, 190, y);
    y += 10;

    // Dados do cliente
    doc.setFontSize(12);
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
    doc.text(`Data da Viagem: ${formatarData(dados.data_vaigem)}`, 20, y);
    doc.text(`Horário: ${dados.horario || ''}`, 80, y);
    doc.text(`Agência: ${dados.agencia || ''}`, 120, y);
    doc.text(`Poltrona: ${dados.poltrona || ''}`, 170, y);

    // Texto adicional e Fone
    y += 15;
    const texto = "Para adiar a passagem o passageiro deverá nos avisar com 02 dias de antecedência, o passageiro tem direito a uma bolsa e uma mala, acima disso será cobrado a parte.";
    const textoQuebrado = doc.splitTextToSize(texto, 170);
    doc.text(textoQuebrado, 20, y);
    y += textoQuebrado.length * 7;
    doc.text(`FONE: (82) 99622-6957`, 20, y);

    // Tabela com Data de Emissão, Valor, Agente ou Agência
    y += 15;
    doc.text(`Data de Emissão: ${formatarData(dados.data_emissao)}`, 20, y);
    doc.text(`Valor: R$ ${dados.valor || ''}`, 80, y);
    doc.text(`Agente/Agência: ${dados.agencia || ''}`, 120, y);

    // Mensagem final
    y += 15;
    doc.setFont('Helvetica', 'bold');
    doc.text('POR FAVOR CHEGAR 1 HORA ANTES DO EMBARQUE * BOA VIAGEM', 105, y, { align: 'center' });

    // Retornar o documento PDF para ser adicionado ao ZIP
    return doc;
}

function formatarData(dataExcel) {
    if (!dataExcel) return '';
    let dataObj;
    if (typeof dataExcel === 'number') {
        // Se for um número, converte do formato Excel para data
        dataObj = new Date((dataExcel - 25569) * 86400 * 1000);
    } else {
        // Se já for uma data, converte para objeto Date
        dataObj = new Date(dataExcel);
    }
    const dia = String(dataObj.getDate()).padStart(2, '0');
    const mes = String(dataObj.getMonth() + 1).padStart(2, '0');
    const ano = dataObj.getFullYear();
    return `${dia}/${mes}/${ano}`;
}
