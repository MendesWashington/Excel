import Express, { Request, Response } from "express";
import Excel from "excel4node";

const app = Express();

app.get("/api/export", (req: Request, res: Response) => {
  try {
    // Criar um novo arquivo Excel
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");

    var HeaderStyle = workbook.createStyle({
      font: {
        bold: true,
        color: "#000000",
        size: 16,
      },
      fill: {
        type: "",
        patternType: "solid",
        bgColor: "#FFFF00",
        fgColor: "#FFFF00",
      },
    });

    // Adicionar os cabeçalhos em negrito
    worksheet.cell(1, 1).string("GERÊNCIA").style(HeaderStyle);
    worksheet.cell(1, 2).string("CARTEIRA").style(HeaderStyle);
    worksheet.cell(1, 3).string("ADQUIRIDO ONLINE").style(HeaderStyle);
    worksheet.cell(1, 4).string("CONSUMIDO ONLINE").style(HeaderStyle);
    worksheet.cell(1, 5).string("SALDO ONLINE").style(HeaderStyle);
    worksheet
      .cell(1, 6)
      .string("PERCENTUAL CONSUMIDO ONLINE")
      .style(HeaderStyle);
    worksheet.cell(1, 7).string("ADQUIRIDO REMOTO").style(HeaderStyle);
    worksheet.cell(1, 8).string("CONSUMIDO REMOTO").style(HeaderStyle);
    worksheet.cell(1, 9).string("SALDO REMOTO").style(HeaderStyle);
    worksheet
      .cell(1, 10)
      .string("PERCENTUAL CONSUMIDO REMOTO")
      .style(HeaderStyle);

    // Primeira coluna
    worksheet.cell(2, 1).string("Gerência 1");
    worksheet.cell(3, 1).string("Gerência 2");
    worksheet.cell(4, 1).string("Gerência 3");
    worksheet.cell(5, 1).string("Gerência 4");

    // Segunda coluna
    worksheet.cell(2, 2).string("Carteira 2");
    worksheet.cell(3, 2).string("Carteira 2");
    worksheet.cell(4, 2).string("Carteira 3");
    worksheet.cell(5, 2).string("Carteira 4");

    // Terceira coluna
    worksheet.cell(2, 3).number(2);
    worksheet.cell(3, 3).number(20);
    worksheet.cell(4, 3).number(60);
    worksheet.cell(5, 3).number(2);

    // Quarta coluna
    worksheet.cell(2, 4).number(0);
    worksheet.cell(3, 4).number(2);
    worksheet.cell(4, 4).number(50);
    worksheet.cell(5, 4).number(0);

    // Quinta coluna
    worksheet.cell(2, 5).number(2);
    worksheet.cell(3, 5).number(18);
    worksheet.cell(4, 5).number(10);
    worksheet.cell(5, 5).number(2);
    // Sexta coluna
    worksheet
      .cell(2, 6)
      .formula("=(D2/C2)*100")
      .style({ numberFormat: "0.00%" });
    worksheet
      .cell(3, 6)
      .formula("=(D3/C3)*100")
      .style({ numberFormat: "0.00%" });
    worksheet
      .cell(4, 6)
      .formula("=(D4/C4)*100")
      .style({ numberFormat: "0.00%" });
    // Sétima coluna
    worksheet.cell(2, 7).number(0);
    worksheet.cell(3, 7).number(30);
    worksheet.cell(4, 7).number(70);
    // Oitava coluna
    worksheet.cell(2, 8).number(0);
    worksheet.cell(3, 8).number(0);
    worksheet.cell(4, 8).number(70);
    // Nona coluna
    worksheet.cell(2, 9).number(0);
    worksheet.cell(3, 9).number(30);
    worksheet.cell(4, 9).number(0);
    worksheet.cell(5, 9).number(0);

    // Décima coluna
    worksheet
      .cell(2, 6)
      .formula("=(H2/G2)*100")
      .style({ numberFormat: "0.00%" });
    worksheet
      .cell(3, 6)
      .formula("=(H3/G3)*100")
      .style({ numberFormat: "0.00%" });
    worksheet
      .cell(4, 6)
      .formula("=(H4/G4)*100")
      .style({ numberFormat: "0.00%" });

    // Definir o cabeçalho de resposta para o tipo de arquivo Excel
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=export.xlsx");

    // Enviar o arquivo Excel como resposta
    workbook.writeToBuffer().then((buffer: string) => {
      let binarybuffer = new Buffer(buffer, "binary");
      res.attachment("filename.xlsx");
      return res.send(binarybuffer);
    });
  } catch (error) {
    return res.status(500).send(error);
  }
});

app.listen(3000, () => {
  console.log("http://localhost:3000");
});
