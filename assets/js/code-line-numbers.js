(function () {
  // Prevent double execution
  if (window.__codeLineNumbersInitialized) return;
  window.__codeLineNumbersInitialized = true;

  const codeBlocks = document.querySelectorAll(
    '.code-block-enhanced.show-lines pre code:not(.lined)'
  );

  codeBlocks.forEach(block => {
    const lines = block.textContent.split('\n');
    const lineCount = lines[lines.length - 1] === ''
      ? lines.length - 1
      : lines.length;

    const lineNumbers = document.createElement('div');
    lineNumbers.className = 'line-numbers';
    lineNumbers.setAttribute('aria-hidden', 'true');

    for (let i = 1; i <= lineCount; i++) {
      const span = document.createElement('span');
      span.textContent = i;
      lineNumbers.appendChild(span);
    }

    const pre = block.parentElement;
    const wrapper = document.createElement('div');
    wrapper.className = 'code-with-lines';

    pre.parentNode.insertBefore(wrapper, pre);
    wrapper.appendChild(lineNumbers);
    wrapper.appendChild(pre);

    block.classList.add('lined');
  });
})();
