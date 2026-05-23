// FluentUI Image Zoom/Lightbox
document.addEventListener('DOMContentLoaded', function() {
  const images = document.querySelectorAll('.image-zoomable');
  let lightbox = null;
  let lightboxImg = null;
  let lightboxCaption = null;

  // Create lightbox overlay on first use
  function createLightbox() {
    if (lightbox) return;

    lightbox = document.createElement('div');
    lightbox.className = 'lightbox';
    lightbox.innerHTML = `
      <div class="lightbox-backdrop"></div>
      <div class="lightbox-content">
        <button class="lightbox-close" aria-label="Close">
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <line x1="18" y1="6" x2="6" y2="18"></line>
            <line x1="6" y1="6" x2="18" y2="18"></line>
          </svg>
        </button>
        <img class="lightbox-image" alt="" />
        <div class="lightbox-caption"></div>
      </div>
    `;
    document.body.appendChild(lightbox);

    lightboxImg = lightbox.querySelector('.lightbox-image');
    lightboxCaption = lightbox.querySelector('.lightbox-caption');

    // Close handlers
    lightbox.querySelector('.lightbox-close').addEventListener('click', closeLightbox);
    lightbox.querySelector('.lightbox-backdrop').addEventListener('click', closeLightbox);
    
    // Keyboard navigation
    document.addEventListener('keydown', function(e) {
      if (lightbox.classList.contains('active')) {
        if (e.key === 'Escape') {
          closeLightbox();
        }
      }
    });
  }

  // Open lightbox
  function openLightbox(imgSrc, imgAlt, caption) {
    createLightbox();
    
    lightboxImg.src = imgSrc;
    lightboxImg.alt = imgAlt;
    
    if (caption) {
      lightboxCaption.textContent = caption;
      lightboxCaption.style.display = 'block';
    } else {
      lightboxCaption.style.display = 'none';
    }

    lightbox.classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  // Close lightbox
  function closeLightbox() {
    if (!lightbox) return;
    
    lightbox.classList.remove('active');
    document.body.style.overflow = '';
    
    // Clear image after animation
    setTimeout(() => {
      lightboxImg.src = '';
    }, 300);
  }

  // Add click handlers to all zoomable images
  images.forEach(img => {
    img.style.cursor = 'zoom-in';
    
    img.addEventListener('click', function() {
      const src = this.getAttribute('data-zoom-src') || this.src;
      const alt = this.alt;
      const figure = this.closest('figure');
      const caption = figure ? figure.querySelector('figcaption')?.textContent : null;
      
      openLightbox(src, alt, caption);
    });
  });
});
