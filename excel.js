let selectedFile;
console.log(window.XLSX);
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

let data = [{
    "name": "jayanth",
    "data": "scd",
    "abc": "sdef"
}]

let permission = 0;

let filial = "AMONTADA    CE"
let enderecoLoja = "AV. GENERAL ALIPIO DOS SANTOS, 916"
let telefoneLoja = "(85) 98131-5140"

let dataa = [];


class Cliente {

    aut = false;
    nome;
    cpf = null;
    bairro;
    estado;
    endereco;
    observacao;
    numero;
    cep;
    complemento;
    cidade;
    celular;
    telefone;
    bloco;

    contas = [];

    carta = ["Carta 1", "Carta 2", "Carta 3"]

    cartaEscolhida;


    getAllVariables() {
        var variables = []

        for (var name in this) {
            if (name != 'cartaEscolhida') {
                variables.push(name)
            }

        }

        return variables;
    }

}

class Conta {

    documento;
    venda;
    data_vencimento;
    data_lancamento;
    diasAtraso;
    valorOriginal;
    juros;
    multa;
    valorRecebido;
    valorReajustado;

    getAllValues() {
        var values = []
        var i = 0;
        Object.values(this).forEach(x => {
            if (i != 0 && i != 3) {
                values.push(x)
            }
            i = i + 1
        })

        return values;
    }

    getAllVariables() {
        var variables = []

        for (var name in this) {
            if (name != 'documento' && name != 'data_lancamento') {
                variables.push(name)
            }



        }

        return variables;
    }


}


document.getElementById('button').addEventListener("click", () => {


    if (permission == 0) {
        XLSX.utils.json_to_sheet(data, 'out.xlsx');
        if (selectedFile) {
            let fileReader = new FileReader();
            fileReader.readAsBinaryString(selectedFile);
            fileReader.onload = (event) => {
                let data = event.target.result;
                let workbook = XLSX.read(data, { type: "binary" });
                console.log(workbook);
                workbook.SheetNames.forEach(sheet => {
                    let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
                    dataa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
                    // document.getElementById("jsondata").innerHTML = JSON.stringify(rowObject, undefined, 4)
                    getData();

                });
            }
        }

        permission = 1;
    }


});


var clientesArray = [];
var contasArray = [];

function getData() {

    console.log(dataa.length)

    for (var i = 0; i < dataa.length; i++) {
        c = new Cliente();
        cc = new Conta();

        c.nome = dataa[i]["Cliente"]

        c.cpf = dataa[i]["Documento"]
        c.bairro = dataa[i]["Bairro"]
        c.endereco = dataa[i]["Endereço"]
        c.numero = dataa[i]["Número"]
        c.complemento = dataa[i]["Complemento"]
        c.observacao = dataa[i]["Complemento"]
        c.cidade = dataa[i]["Cidade"]
        c.telefone = dataa[i]["telefone"]
        c.celular = dataa[i]["celular"]
        c.estado = dataa[i]["Estado"]
        c.cep = dataa[i]["CEP"]


        cc.documento = dataa[i]["Documento"]
        cc.data_lancamento = dataa[i]["Data Lançamento"]
        cc.data_vencimento = dataa[i]["Data Vencimento"]
        cc.valorOriginal = dataa[i]["Valor Original"]
        cc.venda = dataa[i]["Venda Nº"]
        cc.valorReajustado = dataa[i]["Valor Reajustado"]
        cc.multa = dataa[i]["Multa"]
        cc.valorRecebido = dataa[i]["Valor Recebido"]
        cc.diasAtraso = dataa[i]["Dias Em Atraso"]
        cc.juros = dataa[i]["Juros aplicado"]

        //console.log("Aqio - " + cc.valorOriginal + "  " + typeof cc.valorOriginal)

        state = false;

        for (var y = 0; y < clientesArray.length; y++) {
            if (clientesArray[y].cpf == c.cpf) {
                state = true;
            }
        }

        if (!state && typeof c.cpf != 'undefined') {
            clientesArray.push(c)
        }

        if (typeof cc.documento != 'undefined') {
            contasArray.push(cc)
        }

    }


    for (var y = 0; y < contasArray.length; y++) {
        for (var l = 0; l < clientesArray.length; l++) {

            if (contasArray[y].documento == clientesArray[l].cpf) {
                clientesArray[l].contas.push(contasArray[y])
            }

        }
    }


    console.log(contasArray)
    console.log(clientesArray)

    makeTable(clientesArray)

}

function makeTable(listaClientes = []) {



    myTable = document.querySelector("#table")

    let table = document.createElement('table')
    let headerRow = document.createElement('tr')

    cliente = new Cliente();
    variaveis = cliente.getAllVariables();



    variaveis.forEach(nameHeader => {
        let header = document.createElement('th');
        let textNode = document.createTextNode("    " + nameHeader + "    ");

        header.appendChild(textNode)
        headerRow.appendChild(header)
    })


    table.appendChild(headerRow);

    var p = 0;

    listaClientes.forEach((c,index) => {


        let row = document.createElement('tr')

        let cellCheck = document.createElement('td');
        var chk = document.createElement('input');

        chk.setAttribute('class', "check")
        chk.setAttribute('type', 'checkbox');
        chk.setAttribute('id', 'checkBox' + p);

        cellCheck.appendChild(chk);
        row.appendChild(cellCheck)

        let cellNome = document.createElement('td');
        let textNode = document.createTextNode(c.nome);

        cellNome.appendChild(textNode)
        row.appendChild(cellNome)

        let cellCpf = document.createElement('td');
        let textCpf = document.createTextNode(c.cpf);

        cellCpf.appendChild(textCpf)
        row.appendChild(cellCpf)

        let cellBairro = document.createElement('td');
        let textBairro = document.createTextNode(c.bairro);

        cellBairro.appendChild(textBairro)
        row.appendChild(cellBairro)

        let cellEstado = document.createElement('td');
        let textEstado = document.createTextNode(c.estado);

        cellEstado.appendChild(textEstado)
        row.appendChild(cellEstado)

        let cellEndereco = document.createElement('td');
        let textEndereco = document.createTextNode(c.endereco);

        cellEndereco.appendChild(textEndereco)
        row.appendChild(cellEndereco)

        let cellObservaco = document.createElement('td');
        let textObservacao = document.createTextNode(c.observacao);

        cellObservaco.appendChild(textObservacao)
        row.appendChild(cellObservaco)

        let cellNumero = document.createElement('td');
        let textNumero = document.createTextNode(c.numero);

        cellNumero.appendChild(textNumero)
        row.appendChild(cellNumero)

        let cellCep = document.createElement('td');
        let textCep = document.createTextNode(c.cep);

        cellCep.appendChild(textCep)
        row.appendChild(cellCep)


        let cellComplemento = document.createElement('td');
        let textComplemento = document.createTextNode(c.complemento);

        cellComplemento.appendChild(textComplemento)
        row.appendChild(cellComplemento)

        let cellCidade = document.createElement('td');
        let textCidade = document.createTextNode(c.cidade);

        cellCidade.appendChild(textCidade)
        row.appendChild(cellCidade)

        let cellCelular = document.createElement('td');
        let textCelular = document.createTextNode(c.celular);

        cellCelular.appendChild(textCelular)
        row.appendChild(cellCelular)

        let cellTelefone = document.createElement('td');
        let textTelefone = document.createTextNode(c.telefone);

        cellTelefone.appendChild(textTelefone)
        row.appendChild(cellTelefone)

        let cellBloco = document.createElement('td');
        let inputBloco = document.createElement('input')
        inputBloco.setAttribute('class', 'b')
        inputBloco.setAttribute('id', 'bloco'+p)

        cellBloco.appendChild(inputBloco)
        row.appendChild(cellBloco)

        let cellContas = document.createElement('td');
        textConta = "";
        for (var cont = 0; cont < c.contas.length; cont++) {
            textConta = textConta + " Venda : " + c.contas[cont].venda + ", Valor : " + c.contas[cont].valorOriginal;
        }
        let textContas = document.createTextNode(textConta);

        cellContas.appendChild(textContas)
        row.appendChild(cellContas)

        let cellCombo = document.createElement('td');
        var com = document.createElement('select');

        com.setAttribute('id', 'comboBox' + p);
        com.setAttribute('class', 'combo')

        for (var i = 0; i < 3; i++) {
            var op = document.createElement('option')
            op.setAttribute('value', (c.carta[i]))
            op.appendChild(document.createTextNode("carta " + (i + 1)))
            com.appendChild(op)
        }

        cellCombo.appendChild(com)
        row.appendChild(cellCombo)

        table.appendChild(row)
        p = p + 1;
    })


    myTable.appendChild(table)

}



function makeCart() {

    listaClientes = []
    listaOpcaoCarta = []

    document.querySelectorAll('.check').forEach(item => {
        if (item.checked) {

            var comboBox = document.getElementById('comboBox' + item.id.replace(/\D/g, ""))

            cliente = clientesArray[item.id.replace(/\D/g, "")]
            cliente.cartaEscolhida = cliente.carta[comboBox.selectedIndex]

            cliente.bloco = document.getElementById('bloco'+item.id.replace(/\D/g, "")).value;

            listaClientes.push(cliente)

            
        }
    })

    console.log(listaClientes)
    makePdf(listaClientes);

}



function formatDate(date, format) {
    const map = {
        mm: date.getMonth() + 1,
        dd: date.getDate(),
        aa: date.getFullYear().toString().slice(-2),
        aaaa: date.getFullYear()
    }

    return format.replace(/mm|dd|aa|aaaa/gi, matched => map[matched])
}


function makePdf(listaClientes = []) {

    var doc = new jsPDF()

    var first = false;

    listaClientes.forEach(cliente => {

        if (first) {
            doc.addPage();
        }

        first = true;


        doc.setFont("helvetica")
        doc.setFontSize(15)

        doc.text("Ópticas Redenção", 5, 7)

        doc.setFontSize(10)

        doc.text("Filial : " + filial, 5, 15)

        doc.text("______________________________________________________________________________________________________", 5, 20)

        doc.text("Emitido em " + formatDate(new Date(), 'dd/mm/aa' + "\n" + cliente.cartaEscolhida), 175, 10)


        doc.text("Cliente : " + cliente.nome, 5, 28)
        doc.text("CPF : " + cliente.cpf, 100, 28)
        doc.text("Cidade : " + cliente.cidade, 5, 38)
        doc.text("Endereço : " + cliente.endereco, 80, 38)

        if (typeof cliente.numero === `undefined`) {
            cliente.numero = '';
        }
        doc.text("Numero : " + cliente.numero, 145, 38)

        if (typeof cliente.estado === `undefined`) {
            cliente.estado = '';
        }
        doc.text("Estado : " + cliente.estado, 180, 38)

        if (typeof cliente.complemento === `undefined`) {
            cliente.complemento = '';
        }
        doc.text("Complemento : " + cliente.complemento, 5, 48)

        if (typeof cliente.cep === `undefined`) {
            cliente.cep = '';
        }
        doc.text("CEP : " + cliente.cep, 160, 48)

        if (typeof cliente.telefone === `undefined`) {
            cliente.telefone = '';
        }
        doc.text("Telefone : " + cliente.telefone, 5, 58)

        if (typeof cliente.celular === `undefined`) {
            cliente.celular = '';
        }
        doc.text("Celular : " + cliente.celular, 85, 58)

        if (typeof cliente.observacao === `undefined`) {
            cliente.observacao = '';
        }

        doc.text("Carta referente ao(s) bloco(s) : " + cliente.bloco, 5, 68)

        if (cliente.cartaEscolhida == "Carta 1") {
            doc.text("Prezado(a) Cliente", 15, 90)

            doc.text("Uma de nossas constantes preocupações é manter o cliente sempre informado quanto sua situação de crédito até\no momento não constatamos o pagamento do valor de sua(s) parcela(s) Vencida(s) desde " + cliente.contas[0].data_vencimento + " Conforme \ndescrito abaixo", 15, 100)
        } else if (cliente.cartaEscolhida == "Carta 2") {
            doc.text("Prezado(a) Cliente", 15, 90)

            doc.text("Comunicamos pela segunda vez que seu crédito com nossa empresa ainda não foi regularizado, desde " + cliente.contas[0].data_vencimento + "\nconforme descrito abaixo:", 15, 100)
        } else if (cliente.cartaEscolhida == "Carta 3") {
            doc.text("Prezado(a) Cliente", 15, 90)

            doc.text("Visto que pela segunda vez não houve seu comparecimento para resgatar seu crédito com nossa empresa vencido\ndesde " + cliente.contas[0].data_vencimento + " conforme descrito abaixo:", 15, 100)
        }

        contas = []
        for (var i = 0; i < cliente.contas.length; i++) {
            contas.push(cliente.contas[i].getAllValues())
        }

        var totalValor = 0;

        for (var i = 0; i < cliente.contas.length; i++) {
            totalValor = totalValor + parseFloat(cliente.contas[i].valorOriginal)
        }

        contas.push(["", "", "", "", "", "", "", "Total : " + totalValor])

        hears = ['Venda', 'Data Vencimento', 'Atraso(dias)', "Valor Original", "Juros", 'Multa', 'Valor Recebido', 'Total']

        doc.autoTable({
            theme: 'plain',
            styles: {
                fontSize: 6,
            },
            margin: { top: 115 },
            styles: {
                cellWidth: 'wrap'
            },
            head: [hears],
            body: contas
        })

        var altur = 123;

        if (cliente.cartaEscolhida == "Carta 1") {
            doc.text("Aguardamos a presença de V. Sia em nossa empresa para resgatar seu credito, tomando-o assim sempre aceito.\nGostariamos de Salientar a importância da pontualidade do(s) pagamento(s), evitando a cobrança de encargos\nfinanceiros acrescidos em razão do não pagamento. Caso ja tenha tomado providência para a regularização do\nseu credito, desconsidere este aviso, ou em caso de dúvidas entre em contato conosco no endereço ou telefone \ninformado abaixo, para melhores esclarecimentos.", 15, (contas.length * 11) + altur)
        } else if (cliente.cartaEscolhida == "Carta 2") {
            doc.text("Solicitamos assim, que V. Sia compareça em nossa empresa no prazo de 15(quinze) dias após o recebimento desta\ncarta, a fim de que possamos de forma amigável compor acordo referente a vossa dívida. Se não tivermos um retorno\nde vossa parte até o término deste prazo, lamentamos em comunicar-lhe que seremos forçados a enviar os seus\ndados para o S.P.C (Serviço de Proteção ao Crédito), desde que, por sua vez ficará impossibilitado(a) a fazer qualquer\ncrediário em nível nacional. Estaremos a vossa disposição de segunda a sexta das 7h às 17h no endereço informado\nabaixo. Desconsidere este aviso caso já tenha tomado as providências para a regularização do seu crédito.", 15, (contas.length * 11) + altur)
        } else if (cliente.cartaEscolhida == "Carta 3") {
            doc.text("Comunicamos-lhe que seus dados foram enviados para o serviço de proteção ao crédito-SPC e no prazo de 72\n(setenta e duas) horas após o recebimento desta carta; o não comparecimento em nossa empresa para amigavel-\nmente regularizar o seu crédito, serão enviados ao ajuizamento de ação e cobrança, ocasionando-lhe assim grandes\ndanos morais e pessoais com a justiça. Convidamos-lhe para que venha cumprir corretamente seu compromisso\nconosco, após o resgate de suas parcelas, entraremos em acordo e seu crédito novamente será liberado, podendo\nassim evitar transtorno futuros. Solicitamos desconsiderar esse comunicado no caso de já terem sido efetuado o(s)\npagamento(s).", 15, (contas.length * 11) + altur)
        }




        altur = (contas.length * 11) + altur;
        doc.text("  " + filial + " " + formatDate(new Date(), 'dd/mm/aa') + "\n____________________", 85, 30 + altur)
        doc.setFontSize(5)
        doc.text("Departamento Financeiro", 95, 40 + altur);


        doc.setFontSize(8)
        doc.text("Deus seja Louvado", 15, 250)


        doc.setFontSize(10)
        doc.text("______________________________________________________________________________________________________", 5, 283)
        doc.text("     " + filial + "           " + enderecoLoja + "             " + telefoneLoja, 35, 290)

    });

    doc.save(formatDate(new Date(), 'dd/mm/aa') + ".pdf");
}


