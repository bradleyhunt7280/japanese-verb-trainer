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
let furiganaSupport = false;
let currentVerb = null;
let currentForm = null;
let isCurrentQuestionRevealed = false;

const furiganaToggleBtn = document.getElementById("furigana-toggle-btn");

const promptCard = document.getElementById("prompt-card");
const dictionaryRow = document.getElementById("dictionary-row");
const targetRow = document.getElementById("target-row");
const dictionaryEl = document.getElementById("dictionary");
const targetFormEl = document.getElementById("target-form");

const answerCard = document.getElementById("answer-card");
const correctStatusEl = document.getElementById("correct-status");
const answerEl = document.getElementById("answer");

const verbTypeEl = document.getElementById("verb-type");
const ichidanExceptionEl = document.getElementById("ichidan-exception");
const exceptionEl = document.getElementById("exception");

const userAnswerEl = document.getElementById("user-answer");

const revealBtn = document.getElementById("reveal-btn");
const nextBtn = document.getElementById("next-btn");
const revealNextBtn = document.getElementById("reveal-next-btn");

const formTogglesEl = document.getElementById("form-toggles");
const toggleFormsBtn = document.getElementById("toggle-forms-btn");
const settingsContainer = document.getElementById("settings-container");

async function loadExcel() {
  const response = await fetch("verbs.xlsx");
  const arrayBuffer = await response.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  verbs = XLSX.utils.sheet_to_json(sheet).filter((row) => {
    return row["Dictionary"] && String(row["Dictionary"]).trim() !== "";
  });

  console.log("Loaded verbs:", verbs);
  console.log("Number of rows:", verbs.length);

  buildFormToggles();
  generateQuestion({ focusInput: false });
}

function randomItem(array) {
  return array[Math.floor(Math.random() * array.length)];
}

function makeRuby(text, reading) {
  if (!reading) return text;
  return `<ruby>${text}<rt>${reading}</rt></ruby>`;
}

function scrollToElement(element, block = "start", delay = 0) {
  setTimeout(() => {
    element.scrollIntoView({ behavior: "smooth", block });
  }, delay);
}

function focusInput(delay = 0, preventScroll = false) {
  setTimeout(() => {
    if (preventScroll) {
      try {
        userAnswerEl.focus({ preventScroll: true });
      } catch {
        userAnswerEl.focus();
      }
    } else {
      userAnswerEl.focus();
    }
  }, delay);
}

function closeKeyboard() {
  userAnswerEl.blur();

  if (document.activeElement instanceof HTMLElement) {
    document.activeElement.blur();
  }
}

function scrollToDictionaryRow(delay = 0) {
  setTimeout(() => {
    const top = dictionaryRow.offsetTop - 8;

    window.scrollTo({
      top: Math.max(0, top),
      behavior: "smooth"
    });
  }, delay);
}

function displayDictionary(forceFurigana = false) {
  const dictionary = currentVerb["Dictionary"];
  const dictionaryReading = String(currentVerb["Dictionary Reading"] ?? "");

  if (furiganaSupport || forceFurigana) {
    dictionaryEl.innerHTML = makeRuby(dictionary, dictionaryReading);
  } else {
    dictionaryEl.textContent = dictionary;
  }
}

function buildFormToggles() {
  formTogglesEl.innerHTML = "";

  formColumns.forEach((form) => {
    const button = document.createElement("button");
    button.type = "button";
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

function generateQuestion({ focusInput: shouldFocusInput = true } = {}) {
  answerCard.classList.add("hidden");
  isCurrentQuestionRevealed = false;

  let attempts = 0;
  let answerValue = "";

  do {
    currentVerb = randomItem(verbs);
    currentForm = randomItem(activeForms);
    answerValue = String(currentVerb[currentForm] ?? "").trim();
    attempts++;
  } while ((!answerValue || answerValue.toUpperCase() === "NA") && attempts < 100);

  displayDictionary(false);
  targetFormEl.textContent = currentForm;

  userAnswerEl.value = "";

  if (shouldFocusInput && window.innerWidth > 600) {
    focusInput();
  }
}

function revealCurrentAnswer() {
  const rawAnswer = String(currentVerb[currentForm] ?? "");
  const answerReadingColumn = `${currentForm} Reading`;

  const rawAnswerReading = String(currentVerb[answerReadingColumn] ?? "");
  const isException = rawAnswer.includes("*");
  const cleanAnswer = rawAnswer.replace("*", "");

  const userAnswer = userAnswerEl.value.trim();
  const correctAnswerReading = rawAnswerReading.trim();
  const isCorrect = userAnswer === correctAnswerReading;

  isCurrentQuestionRevealed = true;

  // Close the mobile keyboard before scrolling to the reveal block.
  closeKeyboard();
  setTimeout(closeKeyboard, 50);

  // Dictionary always gets furigana after reveal, regardless of setting.
  displayDictionary(true);

  correctStatusEl.textContent = isCorrect ? "Correct ✅" : "Incorrect ❌";
  correctStatusEl.className = isCorrect
    ? "result-status correct"
    : "result-status incorrect";

  answerEl.innerHTML = makeRuby(cleanAnswer, correctAnswerReading);

  verbTypeEl.textContent = currentVerb["Type"];
  ichidanExceptionEl.textContent = currentVerb["Ichidan Exception"];
  exceptionEl.textContent = isException ? "Yes" : "No";

  answerCard.classList.remove("hidden");
  scrollToElement(answerCard, "start", 350);
}

function goToNextQuestion() {
  const isMobile = window.innerWidth <= 600;
  const revealIsVisible = !answerCard.classList.contains("hidden");

  if (!isMobile) {
    generateQuestion({ focusInput: false });
    scrollToElement(promptCard, "start", 100);
    focusInput(250);
    return;
  }

  if (revealIsVisible) {
    // First scroll smoothly from Reveal up to the dictionary row.
    // Do this before hiding Reveal so the page layout does not jump.
    scrollToDictionaryRow(0);

    // Then generate the next question once the scroll is underway.
    setTimeout(() => {
      generateQuestion({ focusInput: false });
      focusInput(100, true);

      // Small correction after mobile keyboard animation.
      scrollToDictionaryRow(550);
    }, 450);
  } else {
    // If Next is pressed from the prompt card, just generate immediately.
    generateQuestion({ focusInput: false });
    scrollToDictionaryRow(50);
    focusInput(350, true);
    scrollToDictionaryRow(800);
  }
}

revealBtn.addEventListener("click", revealCurrentAnswer);
nextBtn.addEventListener("click", goToNextQuestion);
revealNextBtn.addEventListener("click", goToNextQuestion);

userAnswerEl.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    const answer = userAnswerEl.value.trim();

    if (answer.length > 0) {
      event.preventDefault();
      closeKeyboard();

      setTimeout(() => {
        revealCurrentAnswer();
      }, 50);
    }
  }
});

toggleFormsBtn.addEventListener("click", () => {
  settingsContainer.classList.toggle("collapsed");

  if (settingsContainer.classList.contains("collapsed")) {
    toggleFormsBtn.textContent = "Show";
  } else {
    toggleFormsBtn.textContent = "Hide";
  }
});

furiganaToggleBtn.addEventListener("click", () => {
  furiganaSupport = !furiganaSupport;

  furiganaToggleBtn.textContent = furiganaSupport ? "On" : "Off";
  furiganaToggleBtn.classList.toggle("active", furiganaSupport);

  if (currentVerb) {
    displayDictionary(isCurrentQuestionRevealed);
  }
});

loadExcel();