// Ойын файлынан сұрақтарды оқу (қосымшалардағы кестелер ойынның өзімен дәл сәйкес болуы үшін)
const fs = require("fs");
const path = require("path");
const html = fs.readFileSync(path.join(__dirname, "../oiyn/index.html"), "utf8");
const grab = name => {
  const start = html.indexOf(`const ${name} = [`);
  const end = html.indexOf("\n];", start);
  return new Function(`return ${html.slice(start + `const ${name} = `.length, end + 2)}`)();
};
const LEVELS = grab("LEVELS");
const TEST = grab("TEST");
const SURVEY = grab("SURVEY");
const GAME_URL = "https://claude.ai/artifact/NWRfvmZBJcKwFvs5nTSzuU";
module.exports = { LEVELS, TEST, SURVEY, GAME_URL,
  totalQ: LEVELS.reduce((a, l) => a + l.q.length, 0),
  totalLetters: LEVELS.reduce((a, l) => a + l.word.length, 0) };
