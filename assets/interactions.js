// Effet de glissement entre les pages (overlay coloré) + parallaxe des formes au survol de la souris
document.addEventListener("DOMContentLoaded", () => {
  const overlay = document.getElementById("pageTransition");

  // Transition de sortie lors d'un clic sur un lien interne
  document.querySelectorAll("a[href$='.html']").forEach(link => {
    link.addEventListener("click", (e) => {
      const href = link.getAttribute("href");
      if (link.target === "_blank" || e.metaKey || e.ctrlKey) return;
      e.preventDefault();
      overlay.classList.add("leaving");
      setTimeout(() => { window.location.href = href; }, 380);
    });
  });

  // Parallaxe douce des "blobs" décoratifs selon la position de la souris
  const blobs = document.querySelectorAll(".blob");
  if (blobs.length) {
    const hero = document.querySelector(".hero");
    hero.addEventListener("mousemove", (e) => {
      const rect = hero.getBoundingClientRect();
      const x = (e.clientX - rect.left) / rect.width - 0.5;
      const y = (e.clientY - rect.top) / rect.height - 0.5;
      blobs.forEach((blob, i) => {
        const strength = 18 + i * 10;
        blob.style.transform = `translate(${x * strength}px, ${y * strength}px)`;
      });
    });
    hero.addEventListener("mouseleave", () => {
      blobs.forEach(blob => blob.style.transform = "translate(0,0)");
    });
  }

  // Espace admin protégé par code secret (simple, côté navigateur - pas une vraie sécurité)
  const ADMIN_PASSCODE = "tunarco2027";
  const SESSION_KEY = "petits_genies_admin_unlocked";

  document.querySelectorAll("[data-admin-toggle]").forEach(btn => {
    const boxId = btn.getAttribute("data-admin-toggle");
    const box = document.getElementById(boxId);
    if (!box) return;

    function unlock() {
      box.style.display = "block";
      btn.textContent = "🔓 Espace admin (déverrouillé)";
      btn.disabled = true;
    }

    if (sessionStorage.getItem(SESSION_KEY) === "1") {
      unlock();
    }

    btn.addEventListener("click", () => {
      const code = window.prompt("Code admin :");
      if (code === null) return;
      if (code === ADMIN_PASSCODE) {
        sessionStorage.setItem(SESSION_KEY, "1");
        unlock();
      } else {
        alert("Code incorrect.");
      }
    });
  });
});
