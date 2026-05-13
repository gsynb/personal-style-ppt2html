const deck = document.querySelector("#deck");
const slides = Array.from(document.querySelectorAll(".academic-slide"));
let current = 0;

if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
  deck?.setAttribute("data-motion", "none");
}

function showSlide(index) {
  current = Math.max(0, Math.min(slides.length - 1, index));
  slides.forEach((slide, slideIndex) => {
    slide.hidden = slideIndex !== current;
  });
  window.history.replaceState(null, "", `#${current + 1}`);
}

function fromHash() {
  const number = Number.parseInt(window.location.hash.replace("#", ""), 10);
  if (Number.isFinite(number)) {
    showSlide(number - 1);
  }
}

document.addEventListener("keydown", (event) => {
  if (event.key === "ArrowRight" || event.key === "PageDown" || event.key === " ") {
    event.preventDefault();
    showSlide(current + 1);
  }
  if (event.key === "ArrowLeft" || event.key === "PageUp") {
    event.preventDefault();
    showSlide(current - 1);
  }
  if (event.key === "Home") {
    event.preventDefault();
    showSlide(0);
  }
  if (event.key === "End") {
    event.preventDefault();
    showSlide(slides.length - 1);
  }
});

window.addEventListener("hashchange", fromHash);
window.addEventListener("beforeprint", () => {
  slides.forEach((slide) => {
    slide.hidden = false;
  });
});
window.addEventListener("afterprint", () => {
  showSlide(current);
});

if (window.matchMedia("print").matches) {
  slides.forEach((slide) => {
    slide.hidden = false;
  });
} else {
  fromHash();
  showSlide(current);
}

deck?.setAttribute("data-slide-count", String(slides.length));
