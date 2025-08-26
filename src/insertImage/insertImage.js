/* insertImage.js — helpers + init + add caption + validate + wiring */

const WEX = {
  PLACEHOLDER_TITLE: "IMAGE_PLACEHOLDER",
  PLACEHOLDER_MARK: "=== EMPLACEMENT IMAGE ===",
  EXPECTED_CAPTION_PREFIX: "Figure",
  EXACT_CAPTION_OPTIONAL: "", // si non vide -> validation stricte
  REQUIRE_REAL_CAPTION: true,
};

let __busyInit = false;
let __busyCaption = false;
let __busyValidate = false;

function showReport(html) {
  const el = document.getElementById("report");
  if (el) el.innerHTML = html;
  else console.log(html.replace(/<[^>]+>/g, ""));
}
function line(ok, msg) {
  const icon = ok ? "✅" : "❌";
  return `<div>${icon} ${msg}</div>`;
}

async function getOrCreateImagePlaceholder(context) {
  const ccs = context.document.contentControls;
  ccs.load("items");
  await context.sync();

  let cc = ccs.items.find((c) => c.title === WEX.PLACEHOLDER_TITLE);
  if (!cc) {
    const body = context.document.body;
    const p = body.insertParagraph(WEX.PLACEHOLDER_MARK, Word.InsertLocation.end);
    cc = p.insertContentControl();
    cc.title = WEX.PLACEHOLDER_TITLE;
    cc.tag = WEX.PLACEHOLDER_TITLE;
    cc.appearance = "BoundingBox";
  }
  cc.load("id,title,tag");
  await context.sync();
  return cc;
}

// Détecte une vraie légende Word dans le CC : style "Légende/Caption" + champ SEQ Figure/Table/Equation
async function findCaptionParagraphInCC(context, cc) {
  const range = cc.getRange();
  const paras = range.paragraphs;
  paras.load("items");
  await context.sync();

  if (!paras.items || paras.items.length === 0) return null;

  // On charge texte, alignement, style + on prépare l'OOXML
  const ooxmlReqs = [];
  for (const p of paras.items) {
    p.load("text,alignment,style");
    ooxmlReqs.push(p.getOoxml());
  }
  await context.sync(); // nécessaire pour remplir style/text ET ooxmlReqs[*].value

  let best = null;

  for (let i = 0; i < paras.items.length; i++) {
    const p = paras.items[i];
    const txt = (p.text || "").trim();
    const style = (p.style || "").toLowerCase();
    const ooxml = (ooxmlReqs[i].value || "");

    // 1) Style de légende (FR/EN)
    const styleOk = style.includes("légende") || style.includes("caption");

    // 2) Champ SEQ Figure/Table/Equation (vraie légende numérotée)
    const hasSEQ = /\bSEQ\s+(Figure|Table|Equation|Équation)\b/i.test(ooxml);

    // 3) Texte qui ressemble à "Figure 1 : ..."
    const looksLike = isCaptionText(txt);

    // Priorité : (style && SEQ) > (style && looksLike) > (looksLike seul)
    if (styleOk && hasSEQ) {
      best = { para: p, styleOk: true, hasSEQ: true, text: txt };
      break;
    }
    if (!best && styleOk && looksLike) {
      best = { para: p, styleOk: true, hasSEQ: false, text: txt };
    }
    if (!best && looksLike) {
      best = { para: p, styleOk: false, hasSEQ: false, text: txt };
    }
  }

  return best; // null ou { para, styleOk, hasSEQ, text }
}



function isCaptionText(text) {
  if (!text) return false;
  // "Figure 1" ou "Figure 1 : ..." / "Figure 1 - ..." (espaces et ponctuation optionnels)
  const re = /^\s*Figure\s*\d+(\s*[:\-–—]\s*.*)?\s*$/i;
  return re.test(text.trim());
}


async function getNextFigureNumber(context) {
  const paras = context.document.body.paragraphs;
  paras.load("items");
  await context.sync();

  let maxNum = 0;
  for (const p of paras.items) p.load("text");
  await context.sync();

  for (const p of paras.items) {
    const m = (p.text || "").match(/^Figure\s*(\d+)\s*:/i);
    if (m) {
      const n = parseInt(m[1], 10);
      if (!isNaN(n) && n > maxNum) maxNum = n;
    }
  }
  return maxNum + 1;
}

function isCentered(para) {
  return para.alignment === "Centered" || para.alignment === Word.Alignment.centered;
}

// --- UI de confirmation ---
function showClearConfirm() {
  const box = document.getElementById("confirm-clear");
  if (!box) return;
  box.style.display = "block";

  // attache les boutons une seule fois
  if (!window.__CLEAR_CONFIRM_WIRED__) {
    window.__CLEAR_CONFIRM_WIRED__ = true;

    const yes = document.getElementById("btn-clear-yes");
    const no = document.getElementById("btn-clear-no");

    if (yes) yes.onclick = () => {
      hideClearConfirm();
      initDocConfirmed(true);   // -> efface puis initialise
    };
    if (no) no.onclick = () => {
      hideClearConfirm();
      showReport(line(false, "Initialisation annulée (doc non vide)."));
    };
  }
}

function hideClearConfirm() {
  const box = document.getElementById("confirm-clear");
  if (box) box.style.display = "none";
}

// --- corps d'initialisation une fois qu'on sait si on efface ou pas ---
async function initDocConfirmed(withClear) {
  if (__busyInit) return;
  __busyInit = true;
  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      if (withClear) {
        body.clear(); // efface tout
      }

      // (le reste est identique à ton init : titre, consignes, placeholder…)
      const title = body.insertParagraph("Exercice : Image + Légende centrée", Word.InsertLocation.start);
      try { title.styleBuiltIn = Word.BuiltInStyleName.title; } catch (_) { }

      body.insertParagraph("", Word.InsertLocation.end);
      body.insertParagraph(
        "Objectif : Insérez une image à l’emplacement indiqué, ajoutez une légende sous l’image, " +
        "centrez la légende et utilisez un format du type « Figure 1 : … ». ",
        Word.InsertLocation.end
      );

      const h2 = body.insertParagraph("Consignes détaillées :", Word.InsertLocation.end);
      try { h2.styleBuiltIn = Word.BuiltInStyleName.heading2; } catch (_) { }

      [
        "1) Insérer une image à l’emplacement indiqué ci-dessous.",
        "2) Ajouter une légende sous l’image (Références > Insérer une légende, ou via le bouton du complément).",
        "3) Centrer la légende horizontalement.",
        "4) (Optionnel) Inclure un numéro automatique : « Figure 1 : … ».",
      ].forEach((s) => body.insertParagraph(s, Word.InsertLocation.end));

      body.insertParagraph("", Word.InsertLocation.end);

      const cc = await getOrCreateImagePlaceholder(context);
      const ccRange = cc.getRange();
      ccRange.load("text");
      await context.sync();
      if (!ccRange.text || !ccRange.text.includes(WEX.PLACEHOLDER_MARK)) {
        ccRange.insertText(WEX.PLACEHOLDER_MARK, "Replace");
      }

      await context.sync();
      showReport(line(true, "Document initialisé. Emplacement image prêt."));
    });
  } catch (err) {
    console.error(err);
    showReport(line(false, "Erreur pendant l’initialisation : " + err.message));
  } finally {
    __busyInit = false;
  }
}


/* ==============
   Initialisation
   ============== */
async function initDoc() {
  if (__busyInit) return;
  __busyInit = true;
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();

      if (body.text.trim().length > 0) {
        // doc non vide -> on montre la bannière Oui/Non dans le panneau
        __busyInit = false;         // on relâche pour autoriser le clic sur Oui/Non
        showClearConfirm();
        return;
      }
    });

    // doc vide -> on initialise directement (sans effacer)
    if (!__busyInit) return;        // si on a montré la bannière, on s'arrête ici
    __busyInit = false;
    await initDocConfirmed(false);

  } catch (err) {
    console.error(err);
    showReport(line(false, "Erreur pendant l’initialisation : " + err.message));
    __busyInit = false;
  }
}


/* ==========
   Validation
   ========== */
async function validateDoc() {
  if (__busyValidate) return;
  __busyValidate = true;
  try {
    await Word.run(async (context) => {
      let results = [];

      const ccs = context.document.contentControls;
      ccs.load("items");
      await context.sync();

      const cc = ccs.items.find((c) => c.title === WEX.PLACEHOLDER_TITLE);
      if (!cc) {
        results.push(line(false, `Emplacement image introuvable (« ${WEX.PLACEHOLDER_TITLE} »).`));
        showReport(results.join(""));
        return;
      }

      const ccRange = cc.getRange();
      const pics = ccRange.inlinePictures;
      pics.load("items");
      await context.sync();

      const hasImage = pics.items && pics.items.length > 0;
      results.push(line(hasImage, hasImage ? "Image détectée à l’emplacement prévu." : "Aucune image insérée."));

      const cap = await findCaptionParagraphInCC(context, cc);
      if (!cap || !cap.para) {
        results.push(line(false, "Aucune légende détectée sous l’image."));
        showReport(results.join(""));
        return;
      }

      // Centrage
      const centered = isCentered(cap.para);
      results.push(line(centered, centered ? "La légende est centrée." : "La légende n’est pas centrée."));

      // Qualité de la légende (vraie légende Word ?)
      if (cap.styleOk && cap.hasSEQ) {
        results.push(line(true, "Vraie légende Word détectée (style + champ SEQ)."));
      } else if (cap.styleOk && !cap.hasSEQ) {
        results.push(line(false, "Style de légende présent, mais pas de champ SEQ (numérotation)."));
      } else if (!cap.styleOk && cap.text) {
        results.push(line(false, "Texte ressemble à une légende, mais ce n’est pas une légende Word."));
      }

      // Texte conforme ?
      const legendText = (cap.text || "").trim();
      let captionOk = false;
      let captionMsg = "";

      if (WEX.EXACT_CAPTION_OPTIONAL && WEX.EXACT_CAPTION_OPTIONAL.length > 0) {
        captionOk = legendText === WEX.EXACT_CAPTION_OPTIONAL;
        captionMsg = captionOk
          ? `Texte exact : « ${WEX.EXACT_CAPTION_OPTIONAL} ».`
          : `Texte incorrect. Attendu : « ${WEX.EXACT_CAPTION_OPTIONAL} » (trouvé : « ${legendText} »).`;
      } else {
        captionOk = isCaptionText(legendText);
        captionMsg = captionOk
          ? `Format valide (ex. « ${WEX.EXPECTED_CAPTION_PREFIX} 1 : … »).`
          : `Format invalide. Attendu : « ${WEX.EXPECTED_CAPTION_PREFIX} X : … » (trouvé : « ${legendText} »).`;
      }
      results.push(line(captionOk, captionMsg));

      // Réussite globale
      const genuineOk = cap.styleOk && cap.hasSEQ; // vraie légende Word
      const allOk = hasImage && centered && captionOk && (WEX.REQUIRE_REAL_CAPTION ? genuineOk : true);

      results.unshift(`<h3>${allOk ? "✅ Validation réussie" : "⚠️ Validation incomplète"}</h3>`);
      showReport(results.join(""));

    });
  } catch (err) {
    console.error(err);
    showReport(line(false, "Erreur pendant la validation : " + err.message));
  } finally {
    __busyValidate = false;
  }
}

/* =====
   Wiring
   ===== */
Office.onReady(() => {
  // ✅ Empêche les doublons si le script est ré-exécuté (live reload, etc.)
  if (window.__WEX_WIRED__) return;
  window.__WEX_WIRED__ = true;

  const app = document.getElementById("app-body");
  const msg = document.getElementById("sideload-msg");
  if (app && msg) {
    msg.style.display = "none";
    app.style.display = "flex";
  }

  // Utilise onclick (écrase d’anciens handlers) plutôt que addEventListener
  const btnInit = document.getElementById("btn-init");
  const btnCaption = document.getElementById("btn-caption");
  const btnValidate = document.getElementById("btn-validate");

  if (btnInit) btnInit.onclick = initDoc;
  //if (btnCaption) btnCaption.onclick = addCaption;
  if (btnValidate) btnValidate.onclick = validateDoc;
});

