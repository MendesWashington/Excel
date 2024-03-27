import Express, { Request, Response } from "express";
import { Buffer } from "node:buffer";
import Excel from "excel4node";
import { format } from "date-fns";

const app = Express();

app.get("/api/export", (req: Request, res: Response) => {
  try {
    // Criar um novo arquivo Excel
    const workbook = new Excel.Workbook();
    const sumario = workbook.addWorksheet("Sumário");
    const extrato = workbook.addWorksheet("Extrato");

    const HeaderStyle = workbook.createStyle({
      font: {
        bold: true,
        color: "#000000",
        size: 16,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        bgColor: "#FFFF00",
        fgColor: "#FFFF00",
      },
    });

    // Adicionar os cabeçalhos em negrito
    sumario.cell(1, 1).string("GERÊNCIA").style(HeaderStyle);
    sumario.cell(1, 2).string("CARTEIRA").style(HeaderStyle);
    sumario.cell(1, 3).string("ADQUIRIDO ONLINE").style(HeaderStyle);
    sumario.cell(1, 4).string("CONSUMIDO ONLINE").style(HeaderStyle);
    sumario.cell(1, 5).string("SALDO ONLINE").style(HeaderStyle);
    sumario.cell(1, 6).string("PERCENTUAL CONSUMIDO ONLINE").style(HeaderStyle);
    sumario.cell(1, 7).string("ADQUIRIDO REMOTO").style(HeaderStyle);
    sumario.cell(1, 8).string("CONSUMIDO REMOTO").style(HeaderStyle);
    sumario.cell(1, 9).string("SALDO REMOTO").style(HeaderStyle);
    sumario
      .cell(1, 10)
      .string("PERCENTUAL CONSUMIDO REMOTO")
      .style(HeaderStyle);

    // Primeira coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 1).string("Gerência " + i);
    }

    // Segunda coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 2).string("Carteira " + i);
    }

    // Terceira coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 3).number(Math.floor(Math.random() * 100));
    }

    // Quarta coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 4).number(Math.floor(Math.random() * 100));
    }

    // Quinta coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 5).number(Math.floor(Math.random() * 100));
    }

    // Sexta coluna
    for (let i = 2; i <= 10000; i++) {
      sumario
        .cell(i, 6)
        .formula(`=(D${i}/C${i})*100`)
        .style({ numberFormat: "0.00%" });
    }

    // Sétima coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 7).number(Math.floor(Math.random() * 100));
    }
    // Oitava coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 8).number(Math.floor(Math.random() * 100));
    }

    // Nona coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 9).number(Math.floor(Math.random() * 100));
    }

    // Décima coluna
    for (let i = 2; i <= 10000; i++) {
      sumario
        .cell(i, 10)
        .formula(`=(H${i}/G${i})*100`)
        .style({ numberFormat: "0.00%" });
    }

    // Adicionar os cabeçalhos em negrito
    sumario.cell(1, 1).string("GERÊNCIA").style(HeaderStyle);
    sumario.cell(1, 2).string("CARTEIRA").style(HeaderStyle);
    sumario.cell(1, 3).string("ADQUIRIDO ONLINE").style(HeaderStyle);
    sumario.cell(1, 4).string("CONSUMIDO ONLINE").style(HeaderStyle);
    sumario.cell(1, 5).string("SALDO ONLINE").style(HeaderStyle);
    sumario.cell(1, 6).string("PERCENTUAL CONSUMIDO ONLINE").style(HeaderStyle);
    sumario.cell(1, 7).string("ADQUIRIDO REMOTO").style(HeaderStyle);
    sumario.cell(1, 8).string("CONSUMIDO REMOTO").style(HeaderStyle);
    sumario.cell(1, 9).string("SALDO REMOTO").style(HeaderStyle);
    sumario
      .cell(1, 10)
      .string("PERCENTUAL CONSUMIDO REMOTO")
      .style(HeaderStyle);

    // Primeira coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 1).string("Gerência " + i);
    }

    // Segunda coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 2).string("Carteira " + i);
    }

    // Terceira coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 3).number(Math.floor(Math.random() * 100));
    }

    // Quarta coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 4).number(Math.floor(Math.random() * 100));
    }

    // Quinta coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 5).number(Math.floor(Math.random() * 100));
    }

    // Sexta coluna
    for (let i = 2; i <= 10000; i++) {
      sumario
        .cell(i, 6)
        .formula(`=(D${i}/C${i})*100`)
        .style({ numberFormat: "0.00%" });
    }

    // Sétima coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 7).number(Math.floor(Math.random() * 100));
    }
    // Oitava coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 8).number(Math.floor(Math.random() * 100));
    }

    // Nona coluna
    for (let i = 2; i <= 10000; i++) {
      sumario.cell(i, 9).number(Math.floor(Math.random() * 100));
    }

    // Décima coluna
    for (let i = 2; i <= 10000; i++) {
      sumario
        .cell(i, 10)
        .formula(`=(H${i}/G${i})*100`)
        .style({ numberFormat: "0.00%" });
    }
    //Extrato
    // Adicionar os cabeçalhos em negrito
    extrato.cell(1, 1).string("GERÊNCIA").style(HeaderStyle);
    extrato.cell(1, 2).string("CARTEIRA").style(HeaderStyle);
    extrato.cell(1, 3).string("ADQUIRIDO ONLINE").style(HeaderStyle);
    extrato.cell(1, 4).string("CONSUMIDO ONLINE").style(HeaderStyle);
    extrato.cell(1, 5).string("SALDO ONLINE").style(HeaderStyle);
    extrato.cell(1, 6).string("PERCENTUAL CONSUMIDO ONLINE").style(HeaderStyle);
    extrato.cell(1, 7).string("ADQUIRIDO REMOTO").style(HeaderStyle);
    extrato.cell(1, 8).string("CONSUMIDO REMOTO").style(HeaderStyle);
    extrato.cell(1, 9).string("SALDO REMOTO").style(HeaderStyle);
    extrato
      .cell(1, 10)
      .string("PERCENTUAL CONSUMIDO REMOTO")
      .style(HeaderStyle);

    // Primeira coluna
    for (let i = 2; i <= 10000; i++) {
      extrato.cell(i, 1).string("Gerência " + i);
    }

    // Segunda coluna
    for (let i = 2; i <= 10000; i++) {
      extrato.cell(i, 2).string("Carteira " + i);
    }

    // Terceira coluna
    for (let i = 2; i <= 10000; i++) {
      extrato.cell(i, 3).number(Math.floor(Math.random() * 100));
    }

    // Quarta coluna
    for (let i = 2; i <= 10000; i++) {
      extrato.cell(i, 4).number(Math.floor(Math.random() * 100));
    }

    // Quinta coluna
    for (let i = 2; i <= 10000; i++) {
      extrato.cell(i, 5).number(Math.floor(Math.random() * 100));
    }

    // Sexta coluna
    for (let i = 2; i <= 10000; i++) {
      extrato
        .cell(i, 6)
        .formula(`=(D${i}/C${i})*100`)
        .style({ numberFormat: "0.00%" });
    }

    // Sétima coluna
    for (let i = 2; i <= 10000; i++) {
      extrato.cell(i, 7).number(Math.floor(Math.random() * 100));
    }
    // Oitava coluna
    for (let i = 2; i <= 10000; i++) {
      extrato.cell(i, 8).number(Math.floor(Math.random() * 100));
    }

    // Nona coluna
    for (let i = 2; i <= 10000; i++) {
      extrato.cell(i, 9).number(Math.floor(Math.random() * 100));
    }

    // Décima coluna
    for (let i = 2; i <= 10000; i++) {
      extrato
        .cell(i, 10)
        .formula(`=(H${i}/G${i})*100`)
        .style({ numberFormat: "0.00%" });
    }

    // Definir o cabeçalho de resposta para o tipo de arquivo Excel
    const currentDate = new Date();
    const formattedDate = format(currentDate, "yyyy-MM-dd_HH:mm:ss");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=relatorio" + formattedDate + ".xlsx"
    );

    // Enviar o arquivo Excel como resposta
    workbook.writeToBuffer().then((buffer: Buffer) => {
      let binarybuffer = Buffer.alloc(buffer.length, buffer, "binary");
      res.attachment(`relatorio${formattedDate}.xlsx`);
      return res.send(binarybuffer);
    });
  } catch (error) {
    console.log(error);
    return res.status(500).send(error);
  }
});

app.listen(3000, () => {
  console.log("http://localhost:3000");
});
