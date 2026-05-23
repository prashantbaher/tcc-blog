/**
 * Table of Contents with Fluent UI-style active tracking
 */

document.addEventListener('DOMContentLoaded', function() {
  const tocList = document.getElementById('toc-list');
  if (!tocList) return;

  // Generate TOC from headings
  const article = document.querySelector('.docs-article');
  if (!article) return;

  const headings = article.querySelectorAll('h2, h3, h4');
  if (headings.length === 0) return;

  // Build TOC list
  headings.forEach(function(heading, index) {
    // Add ID to heading if it doesn't have one
    if (!heading.id) {
      heading.id = 'heading-' + index;
    }

    const level = heading.tagName.toLowerCase();
    const text = heading.textContent;
    const id = heading.id;

    const li = document.createElement('li');
    const a = document.createElement('a');
    
    a.href = '#' + id;
    a.textContent = text;
    a.setAttribute('data-level', level.replace('h', ''));
    a.classList.add('toc-link');

    li.appendChild(a);
    tocList.appendChild(li);
  });

  // Active link tracking on scroll
  const tocLinks = document.querySelectorAll('.toc-link');
  const headingElements = Array.from(headings);

  function updateActiveLink() {
    const scrollPosition = window.scrollY + 100; // Offset for header

    let currentActiveIndex = -1;

    headingElements.forEach(function(heading, index) {
      const headingTop = heading.offsetTop;
      
      if (scrollPosition >= headingTop) {
        currentActiveIndex = index;
      }
    });

    // Remove all active classes
    tocLinks.forEach(function(link) {
      link.classList.remove('active');
    });

    // Add active class to current section
    if (currentActiveIndex >= 0 && tocLinks[currentActiveIndex]) {
      tocLinks[currentActiveIndex].classList.add('active');
    }
  }

  // Smooth scroll on click
  tocLinks.forEach(function(link) {
    link.addEventListener('click', function(e) {
      e.preventDefault();
      const targetId = this.getAttribute('href').substring(1);
      const targetElement = document.getElementById(targetId);

      if (targetElement) {
        const headerOffset = 80;
        const elementPosition = targetElement.offsetTop;
        const offsetPosition = elementPosition - headerOffset;

        window.scrollTo({
          top: offsetPosition,
          behavior: 'smooth'
        });

        // Update active state immediately
        tocLinks.forEach(function(l) {
          l.classList.remove('active');
        });
        this.classList.add('active');
      }
    });
  });

  // Update on scroll
  window.addEventListener('scroll', updateActiveLink, { passive: true });
  
  // Initial update
  updateActiveLink();

  console.log('TOC initialized with', headings.length, 'headings');
});
