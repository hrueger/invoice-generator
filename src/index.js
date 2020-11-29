import * as path from "path";
import * as fs from "fs";
import XLSX from "xlsx";

export async function cli(args) {
    if (!(args[2] && args[2].endsWith(".csv"))) {
        throw Error("You must specify a .csv file!");
    }
    const filePath = path.join(process.cwd(), args[2]);
    if (!fs.existsSync(filePath)) {
        throw Error("Cannot find the .csv file!");
    }
    const content = fs.readFileSync(filePath).toString();
    const projects = {};
    const fieldIndices = {};

    const lines = content.split("\n").map((l) => l.split(",\"").map((f) => f.replace(/"/g, "").replace(/\r/g, "")));
    const [header] = lines.splice(0, 1);

    for (const [idx, field] of Object.entries(header)) {
        fieldIndices[field] = parseInt(idx, 10);
    }

    const finishedData = [];

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
            total.amount += info.amount;
            total.time += info.duration;
            finishedData.push([project, task, new Date(info.date), new Date(info.duration * 1000 - (60*60*1000)), info.amount]);
        }
    }

    const type = "xlsx";
    if (type == "csv") {
        const csv = finishedData.map((l) => l.join());
        fs.writeFileSync(path.join(process.cwd(), "data.csv"), csv.join("\n"));
    } else if (type == "xlsx") {
        const wb = XLSX.utils.book_new();
        wb.Props = {
            Title: "Invoice",
            Subject: "Invoice",
            Author: "Invoice-Generator",
            CreatedDate: new Date(),
        };
        wb.SheetNames.push("Invoice1");
        var ws = XLSX.utils.aoa_to_sheet(finishedData.reverse());
        wb.Sheets.Invoice1 = ws;
        XLSX.writeFile(wb, "invoice.xlsx");
    }

    console.log(`File written successfully.\nYour total time is ${secondsToTime(total.time)} and you earned ${total.amount} €.`)
}


function timeToSeconds(t) {
    const [hours, minutes, seconds] = t.split(":");
    return (parseInt(hours, 10) * 60 * 60) + (parseInt(minutes, 10) * 60) + parseInt(seconds);
}

function secondsToTime(s) {
    var hours   = Math.floor(s / 3600);
    var minutes = Math.floor((s - (hours * 3600)) / 60);
    var seconds = s - (hours * 3600) - (minutes * 60);

    if (hours   < 10) {hours   = "0"+hours;}
    if (minutes < 10) {minutes = "0"+minutes;}
    if (seconds < 10) {seconds = "0"+seconds;}
    return hours+':'+minutes+':'+seconds;
}