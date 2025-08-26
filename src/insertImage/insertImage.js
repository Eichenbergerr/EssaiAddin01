/* insertImage.js — helpers + init + add caption + validate + wiring */

const WEX = {
  PLACEHOLDER_TITLE: "IMAGE_PLACEHOLDER",
  PLACEHOLDER_MARK: "=== EMPLACEMENT IMAGE ===",
  EXPECTED_CAPTION_PREFIX: "Figure",
  EXACT_CAPTION_OPTIONAL: "", // si non vide -> validation stricte
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

async function getFirstNonEmptyParagraphAfter(context, cc) {
  const afterRange = cc.getRange(Word.RangeLocation.after);
  const paras = afterRange.paragraphs;
  paras.load("items");
  await context.sync();

  if (!paras.items || paras.items.length === 0) return null;
  for (const p of paras.items) p.load("text,alignment");
  await context.sync();

  for (const p of paras.items) {
    if (p.text && p.text.trim().length > 0) return p;
  }
  return null;
}

function isCaptionText(text) {
  if (!text) return false;
  const re = new RegExp(`^${WEX.EXPECTED_CAPTION_PREFIX}\\s*\\d+\\s*:\\s*.+`, "i");
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
        const ok = confirm("Le document contient déjà du texte. Voulez-vous tout effacer ?");
        if (!ok) {
          showReport(line(false, "Initialisation annulée (doc non vide)."));
          return;
        }
        body.clear();
      }

      const title = body.insertParagraph("Exercice : Image + Légende centrée", Word.InsertLocation.start);
      title.styleBuiltIn = Word.BuiltInStyleName.title;
      body.insertParagraph("", Word.InsertLocation.end);
      body.insertParagraph(
        "Objectif : Insérez une image à l’emplacement indiqué, ajoutez une légende sous l’image, " +
        "centrez la légende et utilisez un format du type « Figure 1 : … ». ",
        Word.InsertLocation.end
      );

      const h2 = body.insertParagraph("Consignes détaillées :", Word.InsertLocation.end);
      h2.styleBuiltIn = Word.BuiltInStyleName.heading2;
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

/* ==================
   Ajouter la légende
   ================== */
async function addCaption() {
  if (__busyCaption) return;
  __busyCaption = true;
  try {
    await Word.run(async (context) => {
      // 1) Récupérer l’emplacement
      const ccs = context.document.contentControls;
      ccs.load("items");
      await context.sync();

      const cc = ccs.items.find((c) => c.title === WEX.PLACEHOLDER_TITLE);
      if (!cc) {
        showReport(line(false, `Emplacement image introuvable (« ${WEX.PLACEHOLDER_TITLE} »).`));
        return;
      }

      // 2) Vérifier qu’une image est bien dans le CC
      const ccRange = cc.getRange();
      const pics = ccRange.inlinePictures; // OK: propriété existante sur Range
      pics.load("items");
      await context.sync();

      const hasImage = pics.items && pics.items.length > 0;
      if (!hasImage) {
        showReport(line(false, "Aucune image trouvée dans l’emplacement. Insère d’abord l’image."));
        return;
      }

      // 3) Numéro de figure + description depuis le panneau
      const nextNum = await getNextFigureNumber(context);
      let desc = (document.getElementById("caption-desc")?.value || "").trim();
      if (!desc) desc = "Description";

      const captionText = `${WEX.EXPECTED_CAPTION_PREFIX} ${nextNum} : ${desc}`;

      // 4) S’il existe déjà une légende non vide juste après -> remplacer, sinon insérer après le CC
      let legendPara = await getFirstNonEmptyParagraphAfter(context, cc);

      if (legendPara && isCaptionText(legendPara.text)) {
        legendPara.insertText(captionText, "Replace");
      } else {
        // ✅ Insère directement APRÈS le content control (plus fiable que via afterRange.start)
        legendPara = cc.insertParagraph(captionText, Word.InsertLocation.after);
      }

      // 5) Centrage (utilise l’enum officiel)
      legendPara.alignment = Word.Alignment.centered;

      // (Optionnel) Si tu veux quand même tenter d’appliquer le style intégré quand supporté :
      // if (Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      //   try { legendPara.styleBuiltIn = Word.BuiltInStyleName.caption; } catch (_) {}
      // }

      await context.sync();
      showReport(line(true, `Légende ajoutée : « ${captionText} » (centrée).`));
    });
  } catch (err) {
    console.error(err);
    showReport(line(false, "Erreur pendant l’ajout de la légende : " + err.message));
  } finally {
    __busyCaption = false;
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

      const legendPara = await getFirstNonEmptyParagraphAfter(context, cc);
      if (!legendPara) {
        results.push(line(false, "Aucune légende détectée sous l’image."));
        showReport(results.join(""));
        return;
      }

      const centered = isCentered(legendPara);
      results.push(line(centered, centered ? "La légende est centrée." : "La légende n’est pas centrée."));

      const legendText = (legendPara.text || "").trim();
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

      const allOk = hasImage && centered && (WEX.EXACT_CAPTION_OPTIONAL ? legendText === WEX.EXACT_CAPTION_OPTIONAL : captionOk);
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
  if (btnCaption) btnCaption.onclick = addCaption;
  if (btnValidate) btnValidate.onclick = validateDoc;
});

