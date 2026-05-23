// Floating copy button for code blocks
document.addEventListener("click", async (e) => {
  const btn = e.target.closest(".code-copy-btn-float");
  if (!btn) return;

  const block = btn.closest(".code-block-enhanced");
  if (!block) return;

  const code = block.querySelector("pre code");
  if (!code) return;

  try {
    await navigator.clipboard.writeText(code.innerText);

    btn.classList.add("copied");
    btn.textContent = "Copied";

    setTimeout(() => {
      btn.classList.remove("copied");
      btn.textContent = "Copy";
    }, 1500);
  } catch (err) {
    console.error("Failed to copy code", err);
  }
});
