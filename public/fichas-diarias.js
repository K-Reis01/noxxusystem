function closeBrandToolbar() {
  const button = document.querySelector("#brandMenuButton");
  const toolbar = document.querySelector("#brandToolbar");
  const dim = document.querySelector("#pageDim");
  toolbar.hidden = true;
  dim.hidden = true;
  button.setAttribute("aria-expanded", "false");
}

function toggleBrandToolbar() {
  const button = document.querySelector("#brandMenuButton");
  const toolbar = document.querySelector("#brandToolbar");
  const dim = document.querySelector("#pageDim");
  const shouldOpen = toolbar.hidden;
  toolbar.hidden = !shouldOpen;
  dim.hidden = !shouldOpen;
  button.setAttribute("aria-expanded", String(shouldOpen));
}

document.querySelector("#brandMenuButton").addEventListener("click", (event) => {
  event.stopPropagation();
  toggleBrandToolbar();
});

document.addEventListener("click", (event) => {
  if (!event.target.closest(".brand-menu")) closeBrandToolbar();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeBrandToolbar();
});
