document.addEventListener("DOMContentLoaded", () => {
  const processButton = document.getElementById("processButton");
  processButton.disabled = true;
  processButton.textContent = "Loading...";

  // Dark Mode Toggle
  const darkModeToggle = document.getElementById("darkModeToggle");
  const mediaQuery = window.matchMedia("(prefers-color-scheme: dark)");
  const prefersDark = mediaQuery.matches;
  document.body.classList.toggle("dark", prefersDark);
  if (darkModeToggle) darkModeToggle.checked = prefersDark;

  mediaQuery.addEventListener("change", (e) => {
    document.body.classList.toggle("dark", e.matches);
    if (darkModeToggle) darkModeToggle.checked = e.matches;
  });

  if (darkModeToggle) {
    darkModeToggle.addEventListener("change", () => {
      document.body.classList.toggle("dark");
    });
  }

  // Scroll to Top Button
  const scrollTopBtn = document.getElementById("scrollTopBtn");
  window.addEventListener("scroll", () => {
    scrollTopBtn.classList.toggle("show", window.scrollY > 300);
  });
  scrollTopBtn.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
});

// PyScript Readiness
window.addEventListener("py:all-done", () => {
  const loaderOverlay = document.querySelector(".loader-overlay");
  loaderOverlay.style.display = "none";
  const processButton = document.getElementById("processButton");
  processButton.disabled = false;
  processButton.textContent = "Process and Download";
});

// Config: allow multi-assign (internal)
const ALLOW_MULTI_ASSIGN = false;

// Globals for DnD
let DRAGGED_CARD = null;
let DRAGGED_GROUP = null;

// Utility: hovered card and before/after decision
function getHoverCardInfo(container, clientX, clientY) {
  // Find the element at point and climb to a card within this container
  const el = document.elementFromPoint(clientX, clientY);
  if (el) {
    const card = el.closest(".area-card");
    if (card && container.contains(card)) {
      const rect = card.getBoundingClientRect();
      const placeAfter = clientY > rect.top + rect.height / 2;
      return { card, placeAfter };
    }
  }
  // Fallback: choose last card to append if no direct hover
  const cards = [...container.querySelectorAll(".area-card")];
  if (cards.length) return { card: cards[cards.length - 1], placeAfter: true };
  return { card: null, placeAfter: false };
}

// Form submission
document
  .getElementById("extractForm")
  .addEventListener("submit", async function (event) {
    event.preventDefault();

    const files = document.getElementById("files").files;
    if (files.length === 0) {
      displayMessage("Please upload at least one file.", "error");
      return;
    }

    if (
      typeof window.prepare_files !== "function" ||
      typeof window.finalize_files !== "function"
    ) {
      displayMessage(
        "PyScript is not ready yet. Please wait a moment and try again.",
        "error",
      );
      return;
    }

    const processingOverlay = document.getElementById("processingOverlay");
    const processingModal = document.getElementById("processingModal");
    const processingText = document.getElementById("processingText");

    const toFileInfos = async (files) => {
      const out = [];
      for (let file of files) {
        const arrayBuf = await file.arrayBuffer();
        out.push({
          name: file.name,
          data: Array.from(new Uint8Array(arrayBuf)),
        });
      }
      return out;
    };

    showProcessing(
      processingOverlay,
      processingModal,
      processingText,
      "Processing...",
    );

    let prepareResult;
    try {
      prepareResult = await window.prepare_files(await toFileInfos(files));
    } catch (e) {
      hideProcessing(processingOverlay, processingModal);
      displayMessage("An error occurred during processing.", "error");
      return;
    }

    hideProcessing(processingOverlay, processingModal);

    if (prepareResult.type === "error") {
      displayMessage(prepareResult.message, "error");
      return;
    }

    const proceed = await handleYearWarningIfNeeded(prepareResult);
    if (!proceed) {
      return;
    }

    const grouping = await openGroupingModal(prepareResult.areas);
    if (!grouping) {
      return;
    }

    showProcessing(
      processingOverlay,
      processingModal,
      processingText,
      "Processing...",
    );

    try {
      const result = await window.finalize_files(
        await toFileInfos(files),
        grouping,
        proceed,
      );
      await completeProcessingAnimation(
        processingOverlay,
        processingModal,
        processingText,
        result,
      );
    } catch (e) {
      hideProcessing(processingOverlay, processingModal);
      displayMessage("An error occurred during processing.", "error");
    }
  });

function displayMessage(msg, type) {
  const messageDiv = document.getElementById("message");
  messageDiv.innerHTML = msg;
  messageDiv.className = type;
}

function showProcessing(overlay, modal, textEl, text) {
  const spinner = document.querySelector(".processing-spinner");
  const tick = document.querySelector(".tick-mark");
  tick.style.opacity = "0";
  spinner.style.opacity = "1";
  textEl.textContent = text;
  textEl.style.opacity = "1";
  overlay.style.display = "block";
  modal.style.display = "flex";
}

async function completeProcessingAnimation(overlay, modal, textEl, result) {
  const spinner = document.querySelector(".processing-spinner");
  spinner.style.opacity = "0";
  textEl.style.opacity = "0";
  await new Promise((r) => setTimeout(r, 300));
  const tick = document.querySelector(".tick-mark");
  tick.style.opacity = "1";
  textEl.textContent = "Processing complete.";
  textEl.style.opacity = "1";
  await new Promise((r) => setTimeout(r, 1500));
  overlay.style.display = "none";
  modal.style.display = "none";

  displayMessage(result.message, result.type || "success");
  if (result.buffer) {
    const js_buffer = new Uint8Array(result.buffer);
    const blob = new Blob([js_buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, "0");
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const yyyy = today.getFullYear();
    a.download = `Sales Report ${dd}-${mm}-${yyyy}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  }
}

function hideProcessing(overlay, modal) {
  overlay.style.display = "none";
  modal.style.display = "none";
}

async function handleYearWarningIfNeeded(prepareResult) {
  if (!prepareResult.multi_year) return true;
  return await new Promise((resolve) => {
    const overlay = document.getElementById("yearWarnOverlay");
    const modal = document.getElementById("yearWarnModal");
    yearsText.textContent = prepareResult.years.join(", ");
    overlay.style.display = "block";
    modal.style.display = "flex";
    const yesBtn = document.getElementById("yearWarnYes");
    const noBtn = document.getElementById("yearWarnNo");
    const cleanup = () => {
      overlay.style.display = "none";
      modal.style.display = "none";
      yesBtn.removeEventListener("click", onYes);
      noBtn.removeEventListener("click", onNo);
    };
    const onYes = () => {
      cleanup();
      resolve(true);
    };
    const onNo = () => {
      cleanup();
      resolve(false);
    };
    yesBtn.addEventListener("click", onYes);
    noBtn.addEventListener("click", onNo);
  });
}

async function openGroupingModal(areas) {
  return await new Promise((resolve) => {
    const overlay = document.getElementById("groupOverlay");
    const modal = document.getElementById("groupModal");
    overlay.style.display = "block";
    modal.style.display = "flex";

    const bench = document.getElementById("areaBench");
    bench.innerHTML = "";
    areas.forEach((a) => bench.appendChild(makeAreaCard(a)));

    const groupsWrap = document.getElementById("groupsWrap");
    groupsWrap.innerHTML = "";
    groupsWrap.appendChild(makeGroup());
    groupsWrap.appendChild(makeGroup());
    groupsWrap.appendChild(makeAddGroup());

    attachCardDrop(bench);
    attachGroupDnD(groupsWrap);

    const confirmBtn = document.getElementById("groupConfirm");
    const cancelBtn = document.getElementById("groupCancel");

    const cleanup = () => {
      overlay.style.display = "none";
      modal.style.display = "none";
      confirmBtn.removeEventListener("click", onConfirm);
      cancelBtn.removeEventListener("click", onCancel);
    };

    function onConfirm() {
      const groups = [];
      [...groupsWrap.querySelectorAll(".group-container")].forEach((gc) => {
        const cards = [...gc.querySelectorAll(".area-card")];
        const names = cards.map((c) => c.dataset.name);
        if (names.length > 0) groups.push(names);
      });
      cleanup();
      resolve(groups);
    }
    function onCancel() {
      cleanup();
      resolve(null);
    }

    confirmBtn.addEventListener("click", onConfirm);
    cancelBtn.addEventListener("click", onCancel);
  });
}

function makeAreaCard(name) {
  const card = document.createElement("div");
  card.className = "area-card";
  card.draggable = true;
  card.textContent = name;
  card.dataset.name = name;
  card.addEventListener("dragstart", (e) => {
    e.dataTransfer.setData("text/plain", name);
    e.dataTransfer.effectAllowed = "move";
    card.classList.add("dragging");
    DRAGGED_CARD = card;
  });
  card.addEventListener("dragend", () => {
    card.classList.remove("dragging");
    DRAGGED_CARD = null;
  });
  return card;
}

function renumberGroups() {
  document
    .querySelectorAll("#groupsWrap .group-container .group-title")
    .forEach((el, idx) => {
      el.textContent = `Group ${idx + 1}`;
    });
}

function makeGroup() {
  const container = document.createElement("div");
  container.className = "group-container";
  container.draggable = true;

  const header = document.createElement("div");
  header.className = "group-header";

  const title = document.createElement("span");
  title.className = "group-title";
  title.textContent = "Group";

  const trash = document
    .getElementById("trashTemplate")
    .content.cloneNode(true);
  const trashBox = trash.querySelector(".trash-box");
  trashBox.classList.add("group-trash");

  header.appendChild(title);
  header.appendChild(trashBox);

  const body = document.createElement("div");
  body.className = "group-body";

  container.appendChild(header);
  container.appendChild(body);

  attachCardDrop(body);

  container.addEventListener("dragstart", (ev) => {
    if (DRAGGED_CARD) {
      ev.preventDefault();
      return;
    }
    DRAGGED_GROUP = container;
    if (ev && ev.dataTransfer) ev.dataTransfer.setData("text/plain", "group");
    container.classList.add("dragging-group");
  });
  container.addEventListener("dragend", () => {
    container.classList.remove("dragging-group");
    DRAGGED_GROUP = null;
  });

  trashBox.addEventListener("click", () => {
    const bench = document.getElementById("areaBench");
    if (!ALLOW_MULTI_ASSIGN) {
      [...body.querySelectorAll(".area-card")].forEach((c) =>
        bench.appendChild(c),
      );
    }
    container.remove();
    renumberGroups();
  });

  setTimeout(renumberGroups, 0);

  return container;
}

function makeAddGroup() {
  const add = document.createElement("div");
  add.className = "group-add";
  add.textContent = "Add group";
  add.addEventListener("click", () => {
    const wrap = document.getElementById("groupsWrap");
    wrap.insertBefore(makeGroup(), wrap.lastElementChild);
  });
  return add;
}

function attachCardDrop(container) {
  container.addEventListener("dragover", (e) => {
    if (DRAGGED_GROUP) return;
    if (!DRAGGED_CARD) return;
    e.preventDefault();
    container.classList.add("drag-over");
    // Clear previous markers only in this container
    container
      .querySelectorAll(".area-card.insert-before, .area-card.insert-after")
      .forEach((c) => c.classList.remove("insert-before", "insert-after"));
    const { card, placeAfter } = getHoverCardInfo(
      container,
      e.clientX,
      e.clientY,
    );
    if (card) {
      if (placeAfter) card.classList.add("insert-after");
      else card.classList.add("insert-before");
    }
  });
  container.addEventListener("dragleave", () => {
    container.classList.remove("drag-over");
    container
      .querySelectorAll(".area-card.insert-before, .area-card.insert-after")
      .forEach((c) => c.classList.remove("insert-before", "insert-after"));
  });
  container.addEventListener("drop", (e) => {
    const name = e.dataTransfer.getData("text/plain");
    if (name === "group" || DRAGGED_GROUP) {
      container.classList.remove("drag-over");
      container
        .querySelectorAll(".area-card.insert-before, .area-card.insert-after")
        .forEach((c) => c.classList.remove("insert-before", "insert-after"));
      return;
    }

    e.preventDefault();
    container.classList.remove("drag-over");
    if (!name) return;
    const { card, placeAfter } = getHoverCardInfo(
      container,
      e.clientX,
      e.clientY,
    );
    let nodeToInsert = null;
    if (ALLOW_MULTI_ASSIGN) {
      nodeToInsert = makeAreaCard(name);
    } else if (DRAGGED_CARD) {
      nodeToInsert = DRAGGED_CARD;
    } else {
      // last resort fallback to existing
      nodeToInsert =
        container.querySelector(
          `.area-card[data-name="${CSS.escape(name)}"]`,
        ) || makeAreaCard(name);
    }
    if (card) {
      if (placeAfter)
        card.parentNode.insertBefore(nodeToInsert, card.nextSibling);
      else card.parentNode.insertBefore(nodeToInsert, card);
    } else {
      container.appendChild(nodeToInsert);
    }
    container
      .querySelectorAll(".area-card.insert-before, .area-card.insert-after")
      .forEach((c) => c.classList.remove("insert-before", "insert-after"));
  });
}

function attachGroupDnD(groupsWrap) {
  groupsWrap.addEventListener("dragover", (e) => {
    // Only allow group drag-over handling when a group is actually being dragged.
    // If an area-card is being dragged, do not treat the wrap as a group drop target.
    if (!DRAGGED_GROUP || DRAGGED_CARD) return;
    e.preventDefault();

    // optional: small visual helper for group insertion (mirror of area insert logic)
    groupsWrap
      .querySelectorAll(
        ".group-container.insert-before, .group-container.insert-after",
      )
      .forEach((g) => g.classList.remove("insert-before", "insert-after"));
    const el = document.elementFromPoint(e.clientX, e.clientY);
    const target = el ? el.closest(".group-container") : null;
    if (target && target !== DRAGGED_GROUP) {
      const rect = target.getBoundingClientRect();
      const after = e.clientY > rect.top + rect.height / 2;
      target.classList.add(after ? "insert-after" : "insert-before");
    }
  });
  groupsWrap.addEventListener("drop", (e) => {
    // Ignore if not dragging a group, or if an area is being dragged
    if (!DRAGGED_GROUP || DRAGGED_CARD) return;
    e.preventDefault();

    // Clear previous markers
    groupsWrap
      .querySelectorAll(
        ".group-container.insert-before, .group-container.insert-after",
      )
      .forEach((g) => g.classList.remove("insert-before", "insert-after"));

    // Identify the group under the pointer
    const el = document.elementFromPoint(e.clientX, e.clientY);
    let targetGroup = el ? el.closest(".group-container") : null;
    if (!targetGroup) {
      // Fallback to nearest by Y
      const groups = [
        ...groupsWrap.querySelectorAll(".group-container"),
      ].filter((g) => g !== DRAGGED_GROUP);
      if (groups.length === 0) {
        groupsWrap.appendChild(DRAGGED_GROUP);
        renumberGroups();
        return;
      }
      let minDist = Infinity;
      const y = e.clientY;
      groups.forEach((g) => {
        const rect = g.getBoundingClientRect();
        const mid = rect.top + rect.height / 2;
        const d = Math.abs(y - mid);
        if (d < minDist) {
          minDist = d;
          targetGroup = g;
        }
      });
    }
    if (targetGroup && targetGroup !== DRAGGED_GROUP) {
      const rect = targetGroup.getBoundingClientRect();
      const insertBefore = e.clientY < rect.top + rect.height / 2;
      const currentIndex = [...groupsWrap.children].indexOf(DRAGGED_GROUP);
      const targetIndex = [...groupsWrap.children].indexOf(targetGroup);
      if (
        targetGroup === DRAGGED_GROUP ||
        (insertBefore && currentIndex + 1 === targetIndex) ||
        (!insertBefore && currentIndex === targetIndex + 1)
      ) {
        renumberGroups();
        return;
      }

      if (insertBefore) groupsWrap.insertBefore(DRAGGED_GROUP, targetGroup);
      else
        groupsWrap.insertBefore(DRAGGED_GROUP, targetGroup.nextElementSibling);
      renumberGroups();
    }
  });
}