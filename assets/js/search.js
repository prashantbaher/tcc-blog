/**
 * Optimized Lunr.js search dropdown.
 * - Lazy index loading (on focus/first input)
 * - O(1) doc lookup map
 * - Debounced input
 * - Query result cache
 * - Minimal DOM churn
 */
document.addEventListener("DOMContentLoaded", function () {
  const searchInput = document.getElementById("search-input");
  const resultsBox = document.getElementById("search-results");
  if (!searchInput || !resultsBox) return;
  if (typeof lunr === "undefined") return;

  let searchIndex = null;
  let docsById = Object.create(null);
  let loadingPromise = null;
  let lastRenderKey = "";
  const resultCache = new Map();
  const MAX_CACHE_ENTRIES = 60;
  const MAX_RESULTS = 8;
  const MIN_QUERY_LEN = 2;
  const DEBOUNCE_MS = 120;

  function setVisible(visible) {
    resultsBox.style.display = visible ? "block" : "none";
  }

  function clearResults() {
    resultsBox.textContent = "";
    setVisible(false);
    lastRenderKey = "";
  }

  function normalizeQuery(value) {
    return value.trim().toLowerCase().replace(/\s+/g, " ");
  }

  function safeExcerpt(doc) {
    const excerpt = doc.excerpt || "";
    if (excerpt) return excerpt;
    const content = doc.content || "";
    return content.length > 120 ? content.slice(0, 120) + "..." : content;
  }

  function loadSearchIndex() {
    if (searchIndex) return Promise.resolve();
    if (loadingPromise) return loadingPromise;

    loadingPromise = fetch("/search.json", { cache: "force-cache" })
      .then(function (response) {
        if (!response.ok) throw new Error("search.json load failed");
        return response.json();
      })
      .then(function (data) {
        docsById = Object.create(null);
        for (let i = 0; i < data.length; i++) {
          docsById[String(data[i].id)] = data[i];
        }

        searchIndex = lunr(function () {
          this.ref("id");
          this.field("title", { boost: 10 });
          this.field("content", { boost: 5 });
          this.field("excerpt", { boost: 2 });

          for (let i = 0; i < data.length; i++) {
            this.add(data[i]);
          }
        });
      })
      .catch(function (error) {
        console.error("Failed to initialize search index:", error);
      })
      .finally(function () {
        loadingPromise = null;
      });

    return loadingPromise;
  }

  function runSearch(query) {
    if (!searchIndex) return [];

    if (resultCache.has(query)) {
      return resultCache.get(query);
    }

    let results = [];
    try {
      // Fast exact/standard query first.
      results = searchIndex.search(query);
    } catch (_) {
      // ignore and fallback below
    }

    // Fallback for partial typing.
    if (!results.length) {
      try {
        results = searchIndex.search(query + "*");
      } catch (_) {
        results = [];
      }
    }

    const sliced = results.slice(0, MAX_RESULTS);
    resultCache.set(query, sliced);

    // Keep cache bounded.
    if (resultCache.size > MAX_CACHE_ENTRIES) {
      const oldestKey = resultCache.keys().next().value;
      resultCache.delete(oldestKey);
    }

    return sliced;
  }

  function renderResults(results, query) {
    const renderKey =
      query +
      "|" +
      results
        .map(function (r) {
          return r.ref;
        })
        .join(",");

    if (renderKey === lastRenderKey) return;
    lastRenderKey = renderKey;

    resultsBox.textContent = "";

    if (!results.length) {
      const empty = document.createElement("div");
      empty.className = "search-no-results";
      empty.textContent = "No results found";
      resultsBox.appendChild(empty);
      setVisible(true);
      return;
    }

    const fragment = document.createDocumentFragment();
    for (let i = 0; i < results.length; i++) {
      const doc = docsById[String(results[i].ref)];
      if (!doc) continue;

      const resultItem = document.createElement("a");
      resultItem.href = doc.url || "#";
      resultItem.className = "search-result-item";

      const title = document.createElement("strong");
      title.textContent = doc.title || "Untitled";

      const br = document.createElement("br");

      const excerpt = document.createElement("span");
      excerpt.className = "search-excerpt";
      excerpt.textContent = safeExcerpt(doc);

      resultItem.appendChild(title);
      resultItem.appendChild(br);
      resultItem.appendChild(excerpt);
      fragment.appendChild(resultItem);
    }

    resultsBox.appendChild(fragment);
    setVisible(true);
  }

  // Lazy load search index only when search is actually used.
  searchInput.addEventListener("focus", loadSearchIndex, { once: true });

  let debounceTimer = null;
  searchInput.addEventListener("input", function () {
    if (!searchIndex && !loadingPromise) loadSearchIndex();
    if (debounceTimer) clearTimeout(debounceTimer);

    const query = normalizeQuery(searchInput.value);
    if (query.length < MIN_QUERY_LEN) {
      clearResults();
      return;
    }

    debounceTimer = setTimeout(function () {
      if (!searchIndex) return;
      const results = runSearch(query);
      renderResults(results, query);
    }, DEBOUNCE_MS);
  });

  document.addEventListener("click", function (e) {
    if (!resultsBox.contains(e.target) && e.target !== searchInput) {
      setVisible(false);
    }
  });

  searchInput.addEventListener("keydown", function (e) {
    if (e.key === "Escape") {
      setVisible(false);
      searchInput.blur();
    }
  });
});
