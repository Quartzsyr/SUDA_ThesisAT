const repository = "Quartzsyr/SUDA_ThesisAT";
const countNodes = document.querySelectorAll("[data-download-count]");

function formatNumber(value) {
  return new Intl.NumberFormat("zh-CN").format(value);
}

function clamp(value, minimum, maximum) {
  return Math.min(Math.max(value, minimum), maximum);
}

function smoothstep(value) {
  const progress = clamp(value, 0, 1);
  return progress * progress * (3 - 2 * progress);
}

function mix(start, end, progress) {
  return start + (end - start) * progress;
}

async function loadDownloadCount() {
  try {
    const response = await fetch(`https://api.github.com/repos/${repository}/releases`, {
      headers: { Accept: "application/vnd.github+json" },
    });
    if (!response.ok) throw new Error(`GitHub API responded with ${response.status}`);
    const releases = await response.json();
    const downloads = releases.reduce((total, release) => total + release.assets.reduce(
      (assetTotal, asset) => assetTotal + (asset.download_count || 0), 0
    ), 0);
    countNodes.forEach((node) => { node.textContent = formatNumber(downloads); });
  } catch (error) {
    countNodes.forEach((node) => { node.textContent = "GitHub"; });
  }
}

function setupReveals() {
  const elements = document.querySelectorAll(".reveal");
  const observer = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (!entry.isIntersecting) return;
      entry.target.classList.add("is-visible");
      observer.unobserve(entry.target);
    });
  }, { threshold: 0.16 });
  elements.forEach((element, index) => {
    element.style.transitionDelay = `${Math.min(index % 4, 3) * 70}ms`;
    observer.observe(element);
  });
}

async function setupThesisModel() {
  const canvas = document.querySelector("#thesis-canvas");
  if (!canvas || window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;

  const THREE = await import("https://cdn.jsdelivr.net/npm/three@0.170.0/build/three.module.js");
  const renderer = new THREE.WebGLRenderer({ canvas, alpha: true, antialias: true });
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  renderer.outputColorSpace = THREE.SRGBColorSpace;

  const scene = new THREE.Scene();
  const camera = new THREE.PerspectiveCamera(30, 1, 0.1, 100);
  camera.position.set(0, 0, 9.6);

  const book = new THREE.Group();
  scene.add(book);

  const ivory = new THREE.MeshStandardMaterial({ color: 0xf8f2df, roughness: 0.79, side: THREE.DoubleSide });
  const paperSide = new THREE.MeshStandardMaterial({ color: 0xd8cfb9, roughness: 0.94 });
  const cover = new THREE.MeshStandardMaterial({ color: 0xd9744d, roughness: 0.49 });
  const coverDark = new THREE.MeshStandardMaterial({ color: 0xa74735, roughness: 0.55 });
  const green = new THREE.MeshStandardMaterial({ color: 0x173d33, roughness: 0.7 });
  const gold = new THREE.MeshStandardMaterial({ color: 0xd1b36d, roughness: 0.42, metalness: 0.25 });
  const lineMaterial = new THREE.MeshBasicMaterial({ color: 0x24483d, transparent: true, opacity: 0.63, side: THREE.DoubleSide });

  const pageBlock = new THREE.Mesh(new THREE.BoxGeometry(2.91, 4.02, 0.30), paperSide);
  book.add(pageBlock);

  const backCover = new THREE.Mesh(new THREE.BoxGeometry(3.12, 4.24, 0.10), coverDark);
  backCover.position.z = -0.23;
  book.add(backCover);

  const spine = new THREE.Mesh(new THREE.BoxGeometry(0.15, 4.17, 0.54), coverDark);
  spine.position.set(-1.54, 0, 0);
  book.add(spine);

  const frontAssembly = new THREE.Group();
  frontAssembly.position.set(-1.54, 0, 0);
  const frontCover = new THREE.Mesh(new THREE.BoxGeometry(3.12, 4.24, 0.10), cover);
  frontCover.position.set(1.54, 0, 0.23);
  frontAssembly.add(frontCover);

  const inset = new THREE.Mesh(new THREE.PlaneGeometry(2.50, 3.53), green);
  inset.position.set(1.54, 0.02, 0.286);
  frontAssembly.add(inset);

  const titleMark = new THREE.Mesh(new THREE.RingGeometry(0.18, 0.22, 32), gold);
  titleMark.position.set(1.54, 1.12, 0.291);
  frontAssembly.add(titleMark);

  const titleLineMaterial = new THREE.MeshBasicMaterial({ color: 0xf4e8cc, transparent: true, opacity: 0.9 });
  [[1.05, 0.02], [0.84, -0.23], [0.61, -0.48]].forEach(([width, y]) => {
    const line = new THREE.Mesh(new THREE.PlaneGeometry(width, 0.035), titleLineMaterial);
    line.position.set(1.54, y, 0.292);
    frontAssembly.add(line);
  });
  book.add(frontAssembly);

  const turningPage = new THREE.Group();
  turningPage.position.set(-1.45, 0, 0.17);
  const pageLeaf = new THREE.Mesh(new THREE.PlaneGeometry(2.83, 3.91), ivory);
  pageLeaf.position.x = 1.415;
  turningPage.add(pageLeaf);
  for (let index = 0; index < 7; index += 1) {
    const pageLine = new THREE.Mesh(new THREE.PlaneGeometry(1.75 - (index % 3) * 0.18, 0.026), lineMaterial);
    pageLine.position.set(1.43, 1.1 - index * 0.37, 0.006);
    turningPage.add(pageLine);
  }
  book.add(turningPage);

  const pageEdges = new THREE.Group();
  for (let index = 0; index < 7; index += 1) {
    const edge = new THREE.Mesh(
      new THREE.BoxGeometry(0.025, 3.92, 0.016),
      new THREE.MeshBasicMaterial({ color: 0xbcb29c, transparent: true, opacity: 0.58 })
    );
    edge.position.set(1.27 + index * 0.035, 0, 0.01);
    pageEdges.add(edge);
  }
  book.add(pageEdges);

  const ambient = new THREE.HemisphereLight(0xffeed0, 0x08231d, 2.2);
  const key = new THREE.DirectionalLight(0xffe3b9, 3.6);
  key.position.set(-3, 5, 6);
  const rim = new THREE.PointLight(0xd9744d, 12, 13);
  rim.position.set(3.5, 1, 3);
  scene.add(ambient, key, rim);

  const halo = new THREE.Mesh(
    new THREE.RingGeometry(3.7, 3.73, 96),
    new THREE.MeshBasicMaterial({ color: 0xf2d7a2, transparent: true, opacity: 0.13, side: THREE.DoubleSide })
  );
  halo.rotation.set(0.5, 0.15, -0.34);
  halo.position.z = -1.6;
  scene.add(halo);

  const pointer = new THREE.Vector2();
  let scrollProgress = 0;
  let frameId = 0;

  function resize() {
    const { width, height } = canvas.getBoundingClientRect();
    if (!width || !height) return;
    renderer.setSize(width, height, false);
    camera.aspect = width / height;
    camera.updateProjectionMatrix();
  }

  function updateScroll() {
    const scrollableHeight = Math.max(document.documentElement.scrollHeight - window.innerHeight, 1);
    scrollProgress = clamp(window.scrollY / scrollableHeight, 0, 1);
  }

  function pathAt(progress, compact) {
    const stops = compact
      ? [
          { at: 0, x: 0.45, y: -1.55, scale: 0.60 },
          { at: 0.22, x: 0.86, y: 0.15, scale: 0.26 },
          { at: 0.52, x: 0.86, y: 0.15, scale: 0.27 },
          { at: 0.77, x: 0.78, y: -0.06, scale: 0.24 },
          { at: 1, x: 0.16, y: -0.85, scale: 0.42 },
        ]
      : [
          { at: 0, x: 1.78, y: 0.02, scale: 0.95 },
          { at: 0.22, x: 2.72, y: 0.16, scale: 0.35 },
          { at: 0.52, x: 2.78, y: 0.05, scale: 0.42 },
          { at: 0.77, x: 2.58, y: -0.05, scale: 0.32 },
          { at: 1, x: 1.70, y: -0.55, scale: 0.67 },
        ];
    const end = stops.findIndex((stop) => stop.at >= progress);
    if (end <= 0) return stops[0];
    const start = stops[end - 1];
    const finish = stops[end];
    const local = smoothstep((progress - start.at) / (finish.at - start.at));
    return {
      x: mix(start.x, finish.x, local),
      y: mix(start.y, finish.y, local),
      scale: mix(start.scale, finish.scale, local),
    };
  }

  function render(time) {
    const idle = time * 0.00045;
    const compact = window.innerWidth <= 700;
    const target = pathAt(scrollProgress, compact);
    const openProgress = smoothstep((scrollProgress - 0.07) / 0.23) * (1 - smoothstep((scrollProgress - 0.72) / 0.18));
    const pageProgress = smoothstep((scrollProgress - 0.16) / 0.25) * (1 - smoothstep((scrollProgress - 0.63) / 0.22));

    book.position.x += (target.x + pointer.x * (compact ? 0.025 : 0.08) - book.position.x) * 0.045;
    book.position.y += (target.y + Math.sin(idle * 2.1) * 0.045 - book.position.y) * 0.045;
    book.scale.x += (target.scale - book.scale.x) * 0.06;
    book.scale.y += (target.scale - book.scale.y) * 0.06;
    book.scale.z += (target.scale - book.scale.z) * 0.06;
    book.rotation.x += ((-0.22 + scrollProgress * 0.34 - pointer.y * 0.035) - book.rotation.x) * 0.045;
    book.rotation.y += ((-0.45 + scrollProgress * 1.58 + pointer.x * 0.07) - book.rotation.y) * 0.045;
    book.rotation.z += ((-0.06 + Math.sin(scrollProgress * Math.PI) * 0.1) - book.rotation.z) * 0.045;
    frontAssembly.rotation.y += ((-openProgress * 1.10) - frontAssembly.rotation.y) * 0.07;
    turningPage.rotation.y += ((-0.03 - pageProgress * 0.72) - turningPage.rotation.y) * 0.07;
    halo.rotation.z = -0.34 + scrollProgress * 2.2;
    halo.rotation.x = 0.5 + Math.sin(idle) * 0.04;
    canvas.style.opacity = String(mix(1, 0.57, smoothstep((scrollProgress - 0.14) / 0.18)) + 0.18 * smoothstep((scrollProgress - 0.80) / 0.16));
    renderer.render(scene, camera);
    frameId = requestAnimationFrame(render);
  }

  window.addEventListener("resize", resize, { passive: true });
  window.addEventListener("scroll", updateScroll, { passive: true });
  window.addEventListener("pointermove", (event) => {
    pointer.set((event.clientX / window.innerWidth - 0.5) * 2, (event.clientY / window.innerHeight - 0.5) * 2);
  }, { passive: true });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) cancelAnimationFrame(frameId);
    else frameId = requestAnimationFrame(render);
  });

  resize();
  updateScroll();
  frameId = requestAnimationFrame(render);
}

loadDownloadCount();
setupReveals();
setupThesisModel().catch(() => {
  document.querySelector(".thesis-stage")?.classList.add("three-unavailable");
});
