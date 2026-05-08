let verbs = [];

const formColumns = [
  "Present - Affirmative - Polite",
  "Present - Negative - Polite",
  "Past - Affirmative - Polite",
  "Past - Negative - Polite",
  "Volitional - Polite",
  "Desire",
  "Present - Negative - Plain",
  "Past - Negative - Plain",
  "Past - Affirmative - Plain",
  "Linking"
];

let activeForms = [...formColumns];

const dictionaryEl = document.getElementById("dictionary");
const targetFormEl = document.getElementById("target-form");

const verbTypeEl = document.getElementById("verb-type");
const ichidanExceptionEl = document.getElementById("ichidan-exception");
const answerEl = document.getElementById("answer");
const exceptionEl = document.getElementById("exception");

const answerCard = document.getElementById("answer-card");

const revealBtn = document.getElementById("reveal-btn");
const nextBtn = document.getElementById("next-btn");

const formTogglesEl = document.getElementById("form-toggles");

const userAnswerEl = document.getElementById("user-answer");
const correctStatusEl = document.getElementById("correct-status");

let currentVerb = null;
let currentForm = null;

async function loadExcel() {
  const response = await fetch("verbs.xlsx");
  const arrayBuffer = await response.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  verbs = XLSX.utils.sheet_to_json(sheet);

  console.log("Loaded verbs:", verbs);
  console.log("Number of rows:", verbs.length);

  buildFormToggles();
  generateQuestion();
}

function randomItem(array) {
  return array[Math.floor(Math.random() * array.length)];
}

function buildFormToggles() {
  formTogglesEl.innerHTML = "";

  formColumns.forEach((form) => {
    const button = document.createElement("button");
    button.textContent = form;
    button.classList.add("form-toggle", "active");

    button.addEventListener("click", () => {
      if (activeForms.includes(form)) {
        activeForms = activeForms.filter((activeForm) => activeForm !== form);
        button.classList.remove("active");
      } else {
        activeForms.push(form);
        button.classList.add("active");
      }

      if (activeForms.length === 0) {
        activeForms.push(form);
        button.classList.add("active");
        alert("At least one form must stay active.");
      }
    });

    formTogglesEl.appendChild(button);
  });
}

function generateQuestion() {
  answerCard.classList.add("hidden");

  currentVerb = randomItem(verbs);
  currentForm = randomItem(activeForms);

  dictionaryEl.textContent = currentVerb["Dictionary"];
  targetFormEl.textContent = currentForm;

  userAnswerEl.value = "";
  userAnswerEl.focus();
}

revealBtn.addEventListener("click", () => {
  const rawAnswer = String(currentVerb[currentForm] ?? "");
  const answerReadingColumn = `${currentForm} Reading`;

  const rawAnswerReading = String(currentVerb[answerReadingColumn] ?? "");
  const dictionaryReading = String(currentVerb["Dictionary Reading"] ?? "");

  const isException = rawAnswer.includes("*");
  const cleanAnswer = rawAnswer.replace("*", "");

  const userAnswer = userAnswerEl.value.trim();
  const correctAnswerReading = rawAnswerReading.trim();

  const isCorrect = userAnswer === correctAnswerReading;

  // Add dictionary reading to prompt after reveal
  dictionaryEl.innerHTML = makeRuby(
    currentVerb["Dictionary"],
    dictionaryReading
  );

  // Result block
  correctStatusEl.textContent = isCorrect ? "Correct ✅" : "Incorrect ❌";
  correctStatusEl.className = isCorrect
    ? "result-status correct"
    : "result-status incorrect";

  answerEl.innerHTML = makeRuby(
    cleanAnswer,
    correctAnswerReading
  );

  // Verb info block
  verbTypeEl.textContent = currentVerb["Type"];
  ichidanExceptionEl.textContent = currentVerb["Ichidan Exception"];
  exceptionEl.textContent = isException ? "Yes" : "No";

  answerCard.classList.remove("hidden");
});

nextBtn.addEventListener("click", () => {
  generateQuestion();
});

loadExcel();

function makeRuby(text, reading) {
  if (!reading) return text;

  return `<ruby>${text}<rt>${reading}</rt></ruby>`;
}

const toggleFormsBtn = document.getElementById("toggle-forms-btn");
const formsContainer = document.getElementById("forms-container");

userAnswerEl.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    const answer = userAnswerEl.value.trim();

    if (answer.length > 0) {
      revealBtn.click();
    }
  }
});

toggleFormsBtn.addEventListener("click", () => {
  formsContainer.classList.toggle("collapsed");

  if (formsContainer.classList.contains("collapsed")) {
    toggleFormsBtn.textContent = "Show";
  } else {
    toggleFormsBtn.textContent = "Hide";
  }
});