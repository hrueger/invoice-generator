import * as path from "path";
import * as fs from "fs";
import { Paragraph, Document, Footer, Header, Packer, Table, TableRow, TableCell, WidthType, Border, BorderStyle, AlignmentType, HeadingLevel, TableLayoutType, TextRun, PageNumber } from "docx";

export async function cli(args) {
    if (!(args[2] && args[2].endsWith(".csv"))) {
        throw Error("You must specify a .csv file!");
    }
    const filePath = args[2];//path.join(process.cwd(), args[2]);
    if (!fs.existsSync(filePath)) {
        throw Error("Cannot find the .csv file!");
    }
    const content = fs.readFileSync(filePath).toString();
    const fieldIndices = {};

    const lines = content.split("\n").map((l) => l.split(",\"").map((f) => f.replace(/"/g, "").replace(/\r/g, "")));
    const [header] = lines.splice(0, 1);

    for (const [idx, field] of Object.entries(header)) {
        fieldIndices[field] = parseInt(idx, 10);
    }

    const data = [];

    for (const line of lines) {
        if (line.length < 3) {
            continue;
        }
        if (line.length == 17) {
            console.log(line);
        }
        if (!data[line[fieldIndices.Project]]) {
            data[line[fieldIndices.Project]] = {};
        }
        if (!data[line[fieldIndices.Project]][line[fieldIndices.Description]]) {
            data[line[fieldIndices.Project]][line[fieldIndices.Description]] = {
                date: "",
                duration: 0,
                amount: 0,
            };
        }
        data[line[fieldIndices.Project]][line[fieldIndices.Description]].date = line[fieldIndices["Start Date"]];
        data[line[fieldIndices.Project]][line[fieldIndices.Description]].duration += timeToSeconds(line[fieldIndices["Duration (h)"]]);
        data[line[fieldIndices.Project]][line[fieldIndices.Description]].amount += parseFloat(line[fieldIndices["Billable Amount (EUR)"]]);
    }
    
    let total = {time: 0, amount: 0};
    for (const [project, tasks] of Object.entries(data)) {
        for (const [task, info] of Object.entries(tasks)) {
            info.amount = Math.round(info.amount * 100) / 100;
            total.amount += info.amount;
            total.time += info.duration;
        }
    }
    total.amount = Math.round(total.amount * 100) / 100;

    const options = JSON.parse(fs.readFileSync(path.join(__dirname, "../config.json")).toString());
    
    const noBorderStyle = {
        bottom: {
            size: 1,
            color: "#fff",
            style: BorderStyle.NONE,
        },
        left: {
            size: 1,
            color: "#fff",
            style: BorderStyle.NONE,
        },
        right: {
            size: 1,
            color: "#fff",
            style: BorderStyle.NONE,
        },
        top: {
            size: 1,
            color: "#fff",
            style: BorderStyle.NONE,
        },
    };

    const borderBottomOnly = Object.assign({}, noBorderStyle, {
        bottom: {
            size: 3,
            color: "#000",
            style: BorderStyle.SINGLE,
        },
    });

    const headers = {
        default: new Header({
            children: [
                new Table({
                    rows: [
                        new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph({
                                        text: options.invoiceId,
                                        style: "header"
                                    })],
                                    borders: noBorderStyle,
                                }),
                                new TableCell({
                                    children: [new Paragraph({
                                        text: options.me.name,
                                        style: "header-big",
                                        alignment: AlignmentType.RIGHT
                                    })],
                                    borders: noBorderStyle,
                                }),
                            ]
                        }),
                    ],
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                    },
                    borders: noBorderStyle
                }),
            ],
        }),
    };
    const doc = new Document({
        styles: {
            paragraphStyles: [
                {
                    id: "Standard",
                    run: {
                        font: "Calibri",
                    }
                },
                {
                    id: "Heading",
                    run: {
                        font: "Calibri",
                        size: 50,
                    }
                },
                {
                    id: "header",
                    basedOn: "Standard",
                    run: {
                        color: "#777777",
                        allCaps: true,
                    }
                },
                {
                    id: "header-big",
                    basedOn: "header",
                    run: {
                        size: 30,
                    }
                }
            ],
        }
    });
    const tableWidths = {
        project: 110,
        task: 240,
        date: 60,
        duration: 60,
        amount: 50,
    };
    doc.addSection({
        headers,
        footers: {
            default: new Footer({
                children: [
                    new Table({
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: options.me.name, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: options.me.phone, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: `IBAN ${options.me.account.IBAN}`, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: options.me.address, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: options.me.email, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: `BIC ${options.me.account.BIC}`, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: options.me.city, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: options.me.web, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: options.me.account.bank, style: "Standard" })],
                                        borders: noBorderStyle,
                                    }),
                                ],
                            }),
                        ],
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: noBorderStyle
                    }),
                ],
            }),
        },
        children: [
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: options.customers[options.defaultCustomer].name, style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Rechnung", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: options.invoiceId, style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: options.customers[options.defaultCustomer].company, style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Datum", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: new Date().toLocaleDateString(), style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: options.customers[options.defaultCustomer].address, style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Zeitraum", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "29.10. - 12.12.2020", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: options.customers[options.defaultCustomer].city, style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                        ],
                    }),
                ],
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE,
                },
                borders: noBorderStyle
            }),
            new Paragraph({
                text: "Rechnung",
                style: "Heading",
                spacing: {
                    before: 600,
                    after: 300,
                },
            }),
            new Paragraph({
                text: "Ich möchte Ihnen die folgenden Positionen in Rechnung stellen. Die detaillierte Aufschlüsselung finden Sie auf den nächsten Seiten.",
                style: "Standard",
                spacing: {
                    after: 300,
                },
            }),
            new Table({
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: "Projekt", style: "Standard" })],
                                borders: borderBottomOnly,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Dauer", style: "Standard" })],
                                borders: borderBottomOnly,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Betrag", style: "Standard" })],
                                borders: borderBottomOnly,
                            }),
                        ],
                    }),
                    ...Object.entries(data).map(([project, tasks]) => {
                        tasks = Object.values(tasks);
                        return new TableRow({
                            children: [
                                new TableCell({
                                    children: [new Paragraph({ text: project, style: "Standard" })],
                                    borders: noBorderStyle,
                                }),
                                new TableCell({
                                    children: [new Paragraph({ text: secondsToTime(tasks.reduce((a, b) => a + b.duration, 5)) + " h", style: "Standard" })],
                                    borders: noBorderStyle,
                                }),
                                new TableCell({
                                    children: [new Paragraph({ text: Math.round(tasks.reduce((a, b) => a + b.amount, 0)*100) / 100 + " €", style: "Standard" })],
                                    borders: noBorderStyle,
                                }),
                            ],
                        });
                    }),
                ],
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE,
                },
                borders: noBorderStyle
            }),
            new Paragraph({
                text: `Ich bitte Sie, den Betrag von ${total.amount} € unter Angabe der Rechnungsnummer ${options.invoiceId} auf folgendes Konto zu überweisen:`,
                style: "Standard",
                spacing: {
                    before: 300,
                    after: 300,
                },
            }),
            new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: "Kontoverbindung", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: `IBAN ${options.me.account.IBAN}`, style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: "", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: `BIC ${options.me.account.BIC}`, style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: "", style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: options.me.account.bank, style: "Standard" })],
                                borders: noBorderStyle,
                            }),
                        ],
                    }),
                ],
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE,
                },
                borders: noBorderStyle
            }),
            new Paragraph({
                text: "Mit freundlichen Grüßen",
                style: "Standard",
                spacing: {
                    before: 600,
                    after: 300,
                },
            }),
            new Paragraph({
                text: options.me.name,
                style: "Standard"
            }),
        ],
    });

    doc.addSection({
        headers,
        footers: {
            default: new Footer({
                children: [
                    // @ts-ignore
                    new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [
                            new TextRun({
                                style: "Standard",
                                children: ["Seite ", PageNumber.CURRENT, " von ", PageNumber.TOTAL_PAGES],
                            }),
                        ],
                    })
                ]
            }),
        },
        children: [
            new Table({
                rows: [
                    new TableRow({
                        tableHeader: true,
                        children: [
                            new TableCell({
                                children: [new Paragraph({ text: "Projekt", style: "Standard" })],
                                borders: borderBottomOnly,
                                width: {
                                    size: tableWidths.project,
                                    type: WidthType.DXA,
                                }
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Aufgabe", style: "Standard" })],
                                borders: borderBottomOnly,
                                width: {
                                    size: tableWidths.task,
                                    type: WidthType.DXA,
                                }
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Datum", style: "Standard" })],
                                borders: borderBottomOnly,
                                width: {
                                    size: tableWidths.date,
                                    type: WidthType.DXA,
                                }
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Dauer", style: "Standard" })],
                                borders: borderBottomOnly,
                                width: {
                                    size: tableWidths.duration,
                                    type: WidthType.DXA,
                                }
                            }),
                            new TableCell({
                                children: [new Paragraph({ text: "Betrag", style: "Standard" })],
                                borders: borderBottomOnly,
                                width: {
                                    size: tableWidths.amount,
                                    type: WidthType.DXA,
                                }
                            }),
                        ],
                    }),
                    ...Object.entries(data).flatMap(([project, tasks]) => {
                        return Object.entries(tasks).map(([taskName, info]) => {
                            return new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({ text: project, style: "Standard" })],
                                        borders: noBorderStyle,
                                        width: {
                                            size: tableWidths.project,
                                            type: WidthType.DXA,
                                        }
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: taskName, style: "Standard" })],
                                        borders: noBorderStyle,
                                        width: {
                                            size: tableWidths.task,
                                            type: WidthType.DXA,
                                        }
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: info.date, style: "Standard" })],
                                        borders: noBorderStyle,
                                        width: {
                                            size: tableWidths.date,
                                            type: WidthType.DXA,
                                        }
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({ text: secondsToTime(info.duration) + " h", style: "Standard" })],
                                        borders: noBorderStyle,
                                        width: {
                                            size: tableWidths.duration,
                                            type: WidthType.DXA,
                                        }
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({
                                            text: `${info.amount} €`,
                                            style: "Standard",
                                            alignment: AlignmentType.RIGHT,
                                        })],
                                        borders: noBorderStyle,
                                        width: {
                                            size: tableWidths.amount,
                                            type: WidthType.DXA,
                                        }
                                    }),
                                ],
                            });
                        })
                    }),
                ],
                width: {
                    size: 100,
                    type: WidthType.PERCENTAGE,
                },
                layout: TableLayoutType.FIXED,
                borders: noBorderStyle
            }),
        ],
    });


    Packer.toBuffer(doc).then((buffer) => {
        const filename = `Rechnung_${options.invoiceId}.docx`;
        fs.writeFileSync(filename, buffer);
        var exec = require('child_process').exec;
        exec(`${process.platform == "darwin" ? "open" : process.platform == "win32" ? "start" : "xdg-open"} ${filename}`);
    });

    console.log(`File written successfully.\nYour total time is ${secondsToTime(total.time)} h and you earned ${total.amount} €.`)
}


function timeToSeconds(t) {
    const [hours, minutes, seconds] = t.split(":");
    return (parseInt(hours, 10) * 60 * 60) + (parseInt(minutes, 10) * 60) + parseInt(seconds);
}

function secondsToTime(s) {
    var hours   = Math.floor(s / 3600);
    var minutes = Math.floor((s - (hours * 3600)) / 60);
    var seconds = s - (hours * 3600) - (minutes * 60);

    if (seconds >= 30) {
        minutes += 1;
    }

    if (hours   < 10) {hours   = "0"+hours;}
    if (minutes < 10) {minutes = "0"+minutes;}
    return hours+':'+minutes;
}