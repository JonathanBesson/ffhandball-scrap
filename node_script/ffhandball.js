import puppeteer from "puppeteer";
import minimist from "minimist";
import chalk from "chalk";
import fs from "fs";
import ics from "ics";
import xl from "excel4node";
import ftp from "basic-ftp";

const args = minimist(process.argv.slice(2));

const scrap = async (url, teamName) => {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();

  await page.setViewport({
    width: 800,
    height: 600,
  });

  await page.goto(url);

  await page.waitForSelector(".m-standings__title");
  const title = await page.evaluate(() => {
    return document.querySelector(".m-standings__title").textContent;
  });

  await page.waitForSelector(".s-fixtures-table");

  // Extract the results from the page.
  const events = await page.evaluate(() => {
    const scrapData = () => {
      const teams = [...document.querySelectorAll(".s-fixtures-table-cell-row .s-fixtures-table-cell-row__name")].map((team) => team.textContent);
      const dates = [...document.querySelectorAll(".s-fixtures-table-cell-row__date")].map((date) => date.textContent);
      const locations = [...document.querySelectorAll(".s-fixtures-table-cell-list")].map((info) => info.textContent);
      const matchs = [];
      for (id in dates) {
        matchs.push({
          dates: dates[id],
          teams: [teams[id * 2], teams[id * 2 + 1]],
          locations: locations[id * 2],
        });
      }

      return matchs.filter((match) => match.teams.join("").includes("BUGUE ATHLETIQUE CLUB HANDBALL"));
    };

    const dates = [...document.querySelectorAll(".s-fixtures-calendar-date .s-fixtures-calendar-day")];
    if (dates.length >= 1) {
      return dates.map((link) => {
        link.click();
        return [...scrapData()];
      });
    }

    return scrapData();
  });

  // Close browser.
  await browser.close();

  // Print all the files.
  return {
    title,
    events,
  };
};

const flatenizeEvents = async (events) => {
  let flat = [];
  const flatenize = (item) => {
    if (typeof item === "object" && !Array.isArray(item)) {
      flat.push(item);
      return item;
    }

    return item.map((child) => flatenize(child));
  };
  flatenize(events);
  return flat;
};

const makeICSData = async (events) => {
  const icsEvents = events.map((event) => {
    let [dayText, dayNumber, monthText, time] = event.dates.split(" ");
    const year = new Date().getFullYear();
    const months = {
      janvier: 1,
      février: 2,
      mars: 3,
      avril: 4,
      mai: 5,
      juin: 6,
      juillet: 7,
      aout: 8,
      septembre: 9,
      octobre: 10,
      novembre: 11,
      décembre: 12,
    };
    const [hours, minutes] = time.split("H");

    return {
      title: "Match BAC HB contre " + event.teams[0],
      start: [Number(year), Number(months[monthText]), Number(dayNumber), Number(hours), Number(minutes)],
      duration: { hours: 2 },
      description: event.teams.join(" vs "),
      location: event.locations,
    };
  });

  return ics.createEvents(icsEvents);
};

const writeICSFile = async (filename, value) => {
  fs.writeFileSync(`${filename}.ics`, value);
  console.log(chalk.green(`
  Fichier ${filename}.ics enregistré avec succès !
  `));
};

const writeCSVFile = async (filename, sheetname, events) => {
  // Create a new instance of a Workbook class
  const wb = new xl.Workbook();

  // Add Worksheets to the workbook
  const ws = wb.addWorksheet(sheetname);

  // Create a reusable style
  const styleHeader = wb.createStyle({
    font: {
      color: "#000000",
      size: 14,
      bold: true,
    },
    alignment: {
      horizontal: "center",
    },
    fill: {
      type: "pattern",
      bgColor: "#bbbbbb",
    },
  });
  const style = wb.createStyle({
    font: {
      color: "#333333",
      size: 12,
    },
  });

  // set widths
  ws.column(1).setWidth(25);
  ws.column(2).setWidth(80);
  ws.column(3).setWidth(80);
  ws.column(4).setWidth(80);

  const offset = 2;
  const flat = await flatenizeEvents(events);
  if (flat.length < 1) {
    console.log(chalk.red(`
    Aucun évènements à ajouter au Excel.
    `));
    return;
  }

  // header
  ws.cell(1, 1).string("Date").style(styleHeader);
  ws.cell(1, 2).string("Équipe 1").style(styleHeader);
  ws.cell(1, 3).string("Équipe 2").style(styleHeader);
  ws.cell(1, 4).string("Terrain").style(styleHeader);

  flat.forEach((event, id) => {
    ws.cell(id + offset, 1)
      .string(event.dates)
      .style(style);
    ws.cell(id + offset, 2)
      .string(event.teams[0])
      .style(style);
    ws.cell(id + offset, 3)
      .string(event.teams[1])
      .style(style);
    ws.cell(id + offset, 4)
      .string(event.locations)
      .style(style);
  });

  wb.write(`${filename}.xlsx`);
};

const commitToServeur = async (localPath, distantPath) => {
  const client = new ftp.Client();
  client.ftp.verbose = false;
  try {
    await client.access({
      host: "ftp.cluster021.hosting.ovh.net",
      user: "jonathanze",
      password: "f53gMEsqbq1AGqCm0Q0914FxoiN3PC",
      secure: false,
    });
    await client.uploadFrom(localPath, distantPath);
  } catch (err) {
    console.log(err);
  }
  client.close();
};

const run = async () => {
  console.log(
    chalk.yellow(`
  scrap starting at : ${new Date().toString()}
  url : ${args.url}
  team : ${args.team}
  script : node node_script/ffhandball.js --url ${args.url} --team ${args.team}
  `)
  );
  const { events, title } = await scrap(args.url, args.team);

  // cleaning 
  const flat = await flatenizeEvents(events);
  const bookedEvents = flat.filter((event) => event.dates !== "—");

  if (flat.length < 1 || !bookedEvents.length >= 1) {
    console.log(chalk.red(`
    Aucun évènements prévus ou programmés.
    `));
    process.exit(0); // terminer le processus avec un code de retour de 0 (succès)
  }

  console.log(events, title);
  const { error, value } = await makeICSData(bookedEvents);

  await writeICSFile(title, value);
  await commitToServeur(`${title}.ics`, `/me/bac_handball/${title}.ics`);

  await writeCSVFile(title, `Saison ${new Date().getFullYear()}`, bookedEvents);
  await commitToServeur(`${title}.xlsx`, `/me/bac_handball/${title}.xlsx`);

  console.log(chalk.yellow(`
  scrap ended at : ${new Date().toString()}
  `));
  process.exit(0); // terminer le processus avec un code de retour de 0 (succès)
};

run();
