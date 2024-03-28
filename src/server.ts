import Express, { Request, Response } from "express";
import { Buffer } from "node:buffer";
import Excel from "excel4node";
import { format } from "date-fns";
import path from "path";

const app = Express();

app.get("/api/export", (req: Request, res: Response) => {
  try {
    const workbook = new Excel.Workbook();
    const sumario = workbook.addWorksheet("Sumário");
    const extrato = workbook.addWorksheet("Extrato");

    const headerStyle = workbook.createStyle({
      font: {
        bold: true,
        color: "#000000",
        size: 16,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        bgColor: "#b6b6b8",
        fgColor: "#b6b6b8",
      },
      alignment: {
        vertical: "center",
        horizontal: "left",
      },
    });

    const titleStyle = workbook.createStyle({
      font: {
        bold: true,
        color: "#000000",
        size: 24,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        bgColor: "#b6b6b8",
        fgColor: "#b6b6b8",
      },
      alignment: {
        vertical: "center",
        horizontal: "center",
      },
    });

    const borderStyle: any = {
      top: { style: "thin", color: "#949494" },
      bottom: { style: "thin", color: "#949494" },
      left: { style: "thin", color: "#949494" },
      right: { style: "thin", color: "#949494" },
    };

    const imagePath = path.join(__dirname, "./images/logo.png");

    const imagePosition: any = {
      type: "oneCellAnchor",
      from: { col: 1, colOff: "30cm", row: 1, rowOff: "1cm" },
      to: { col: 10, colOff: "2cm", row: 1, rowOff: "2cm" },
    };

    // SUMÁRIO =========================================================

    sumario.addImage({
      path: imagePath,
      type: "picture",
      position: imagePosition,
    });

    sumario.row(1).setHeight(80);

    sumario.cell(2, 1).string("SUMÁRIO").style(titleStyle);

    sumario.row(2).setHeight(50);

    // params: linha inicial, coluna inicial, linha final, coluna final, isMerged
    sumario.cell(2, 1, 2, 10, true);

    const sumarioHeader = [
      "GERÊNCIA",
      "CARTEIRA",
      "ADQUIRIDO ONLINE",
      "CONSUMIDO ONLINE",
      "SALDO ONLINE",
      "PERCENTUAL CONSUMIDO ONLINE",
      "ADQUIRIDO REMOTO",
      "CONSUMIDO REMOTO",
      "SALDO REMOTO",
      "PERCENTUAL CONSUMIDO REMOTO",
    ];

    sumarioHeader.forEach((header, index) => {
      sumario
        .cell(3, index + 1)
        .string(header)
        .style(headerStyle);
      sumario.column(index + 1).setWidth(header.length + 15);
    });

    for (let row = 4; row <= 10000; row++) {
      for (let col = 1; col <= 10; col++) {
        const cell = sumario.cell(row, col);
        cell.style({ border: borderStyle });
      }
      sumario.cell(row, 1).string("Gerência " + row);
      sumario.cell(row, 2).string("Carteira " + row);
      sumario.cell(row, 3).number(Math.floor(Math.random() * 100));
      sumario.cell(row, 4).number(Math.floor(Math.random() * 100));
      sumario
        .cell(row, 5)
        .formula(`=C${row}-D${row}`)
        .style({ numberFormat: "0.00" });
      sumario
        .cell(row, 6)
        .formula(`=D${row}/C${row}`)
        .style({ numberFormat: "0.00%" });
      sumario.cell(row, 7).number(Math.floor(Math.random() * 100));
      sumario.cell(row, 8).number(Math.floor(Math.random() * 100));
      sumario
        .cell(row, 9)
        .formula(`=G${row}-H${row}`)
        .style({ numberFormat: "0.00" });
      sumario
        .cell(row, 10)
        .formula(`=H${row}/G${row}`)
        .style({ numberFormat: "0.00%" });
    }

    // EXTRATO =========================================================

    extrato.addImage({
      path: imagePath,
      type: "picture",
      position: imagePosition,
    });

    extrato.row(1).setHeight(80);

    extrato.cell(2, 1).string("EXTRATO").style(titleStyle);

    extrato.row(2).setHeight(50);

    extrato.cell(2, 1, 2, 14, true);

    const extratoHeader = [
      "CÓDIGO ECG",
      "CÓDIGO CLIENTE",
      "DESCRIÇÃO",
      "GERÊNCIA",
      "CARTEIRA",
      "ON/REMOTA",
      "MODERADO",
      "MODERADOR",
      "OWNER CLIENTE",
      "OWNER ECG",
      "ATIVIDADES",
      "DATA",
      "CRÉDITOS",
      "METODOLOGIA DE PESQUISA",
    ];

    extratoHeader.forEach((header, index) => {
      extrato
        .cell(3, index + 1)
        .string(header)
        .style(headerStyle);
      extrato.column(index + 1).setWidth(header.length + 15);
    });

    for (let row = 4; row <= 1000; row++) {
      for (let col = 1; col <= 14; col++) {
        const cell = extrato.cell(row, col);
        cell.style({ border: borderStyle });
      }
      extrato.cell(row, 1).string("Código ECG");
      extrato.cell(row, 2).string("Código Cliente");
      extrato.cell(row, 3).string("Descrição");
      extrato.cell(row, 4).string("Carteira" + row);
      extrato.cell(row, 5).string("Gerência" + row);
      extrato.cell(row, 6).string("Online ou Remota");
      extrato.cell(row, 7).string("Moderação");
      extrato.cell(row, 8).string("Adelina");
      extrato.cell(row, 9).string("Alan");
      extrato.cell(row, 10).string("Bete");
      extrato.cell(row, 11).string("Atividade");
      extrato.cell(row, 12).string("28/03/24");
      extrato.cell(row, 13).number(Math.floor(Math.random() * 100));
      extrato.cell(row, 14).string("Exploratória");
    }

    // EXTRATO =========================================================

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
