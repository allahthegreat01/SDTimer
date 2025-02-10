/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none"; // Begrüßungsnachricht ausblenden
    document.getElementById("app-body").style.display = "flex"; // Hauptinhalt anzeigen
    document.getElementById("start-button").addEventListener("click", startTimer);
    document.getElementById("reset-button").addEventListener("click", resetTimer);
  }
});

let timerInterval = null; // Speichert die Timer-ID
let remainingTime = 0; // Speichert die verbleibende Zeit
const audio = new Audio("assets/alarm.mp3"); // Pfad zur Audiodatei

function startTimer() {
  const hours = parseInt(document.getElementById("input-hours").value) || 0;
  const minutes = parseInt(document.getElementById("input-minutes").value) || 0;
  const seconds = parseInt(document.getElementById("input-seconds").value) || 0;

  remainingTime = (hours * 3600 + minutes * 60 + seconds) * 1000;

  if (remainingTime <= 0) {
    alert("Bitte eine gültige Zeit eingeben!");
    return;
  }

  // Falls ein Timer läuft, beende ihn
  if (timerInterval) {
    clearInterval(timerInterval);
    timerInterval = null;
  }

  // Sofortige Anzeige aktualisieren
  updateDisplay(remainingTime);

  const endTime = Date.now() + remainingTime;

  // Starte neuen Timer
  timerInterval = setInterval(() => {
    const now = Date.now();
    remainingTime = endTime - now;

    if (remainingTime <= 0) {
      clearInterval(timerInterval);
      timerInterval = null;
      updateDisplay(0);
      playSound();
      return;
    }

    updateDisplay(remainingTime);
  }, 1000);
}

function resetTimer() {
  audio.pause();
  // Stoppe den Timer
  if (timerInterval) {
    clearInterval(timerInterval);
    timerInterval = null;
  }

  remainingTime = 0;

  // Anzeige zurücksetzen
  document.getElementById("timer-display").textContent = "00:00:00";

  // Eingabefelder zurücksetzen
  document.getElementById("input-hours").value = "0";
  document.getElementById("input-minutes").value = "0";
  document.getElementById("input-seconds").value = "0";
}

function updateDisplay(time) {
  const hours = Math.floor(time / 3600000);
  const minutes = Math.floor((time % 3600000) / 60000);
  let seconds = Math.round((time % 60000) / 1000);

  if (seconds === 60) {
    seconds = 0;
    minutes++;
  }

  const formattedTime = `${hours.toString().padStart(2, "0")}:${minutes.toString().padStart(2, "0")}:${seconds.toString().padStart(2, "0")}`;

  // Anzeige im Add-in aktualisieren
  document.getElementById("timer-display").textContent = formattedTime;
}

function playSound() {
  audio.play();
}
