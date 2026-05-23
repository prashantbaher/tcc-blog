// FluentUI Tabs functionality
document.addEventListener('DOMContentLoaded', function() {
  const tabContainers = document.querySelectorAll('.tabs-container');
  
  tabContainers.forEach(container => {
    const buttons = container.querySelectorAll('.tab-button');
    const panels = container.querySelectorAll('.tab-panel');
    
    buttons.forEach((button, index) => {
      button.addEventListener('click', () => {
        // Remove active from all buttons and panels
        buttons.forEach(btn => {
          btn.classList.remove('active');
          btn.setAttribute('aria-selected', 'false');
        });
        panels.forEach(panel => panel.classList.remove('active'));
        
        // Add active to clicked button and corresponding panel
        button.classList.add('active');
        button.setAttribute('aria-selected', 'true');
        panels[index].classList.add('active');
      });
      
      // Keyboard navigation
      button.addEventListener('keydown', (e) => {
        let newIndex = index;
        
        if (e.key === 'ArrowRight') {
          newIndex = (index + 1) % buttons.length;
          e.preventDefault();
        } else if (e.key === 'ArrowLeft') {
          newIndex = (index - 1 + buttons.length) % buttons.length;
          e.preventDefault();
        }
        
        if (newIndex !== index) {
          buttons[newIndex].click();
          buttons[newIndex].focus();
        }
      });
    });
  });
});
