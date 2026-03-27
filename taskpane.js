/* =========================================================
   taskpane.js  —  RTF 智能標記系統 · Word 增益集核心邏輯
   使用 Office.js Word API 直接操作文件格式
   ========================================================= */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById('btnApply').disabled = false;
  }
});

// ── 詞庫 ──────────────────────────────────────────────────
const BUILT_IN_IMPORTANT = '玄空師父,關聖帝君,玉皇上帝,聖神仙佛,救劫經,弟子規,傳世明訓,四句箴言,志業體,台北本宮,北投分宮,三峽分宮,社會大學,待人處事,盡心盡力,腳踏實地,正己化人,克己利人,萬般由心,福由心造,禍福無門,惟人自召,善惡報應,如影隨形,舉頭三尺,神明鑒察,將心比心,設身處地,捨身取義,發光發熱,正向思考,負面思考,自怨自艾,怨天尤人,憤世嫉俗,風調雨順,國泰民安,天地清寧,金玉滿堂,排除萬難,勇往直前,始終如一,有始有終,善始善終,了凡四訓,岳武穆王,司命真君,問心書院,舉世聞名,萬代傳香,永垂不朽,循循善誘,引迷入醒,撥亂反正,聞聲救苦,互相扶持,共同奮鬥,無怨無悔,甘之如飴,修德造命,飲水思源,慎終追遠,民胞物與,敬天畏地,尊重自然,環境保護,永續發展,一心向善,一念之誠,萬事具足,安身立命,瞬息萬變,不生不滅,色即是空,空即是色,六根清淨,慈航普度,救苦求難,大發慈悲,德澤廣被,恩光普照,平安進道,福慧增長,萬事如意,闔家平安,身心健康,守口如瓶,不欺暗室,赤子之心,惻隱之心,功德圓滿,萬德莊嚴,黎民百姓,春季拜斗,秋季拜斗,聖誕典禮,祝壽儀典,改造命運,萬古長存,福慧相隨,合宅榮昌,諸事吉祥,新春愉快,恩主公,呂恩主,志工,聖人,同仁,道德,五倫,八德,五常,仁義,禮義,忠孝,廉恥,忠義,孝順,慈悲,同理心,誠信,正信,謙遜,忍辱,精進,禪定,智慧,節操,正直,良心,天理,誦經,讀經,宣講,解籤,祭解,祈福,功德,願力,善願,善行,斷惡修善,因果,福報,感應,懺悔,覺悟,超度,回向,列聖寶經,明聖經,大洞經,醒心經,行天宮,行修宮,道場,當責,隨緣,知足,惜福,平安,吉祥,圓滿,和諧,安定,祥和,奉獻,歡喜,逆境,順境,抄經,收驚,掩魂,籤詩,勸善,救苦,慈心,悲心,至誠,真誠,誠懇,踏實,積極,勤奮,努力,堅持,恆心,毅力,挫折,挑戰,磨練,考驗,覺察,反省,檢討,改過,提升,健全,人格,氣質,涵養,學問,本分,職責,凝聚力,向心力,尊重,包容,體諒,感恩,回饋,報恩,守法,規矩,紀律,公平,正義,廉潔,知恥,氣節,剛正,坦蕩,無私,貪瞋癡,感化,帶動,傳遞,循環,善循環,樂觀,悲觀,堅強,謙虛,大度,慷慨,施予,利他,自度,度人,成全,成熟,執著,貪欲,瞋恨,愚癡,嫉妒,驕傲,傲慢,怠惰,欺騙,妄語,惡口,爭取,陷害,毀謗,猜疑,忠厚,寬洪,儉約,孤兒,寡婦,老幼,弱勢,貧窮,祥瑞,修真,得道,成聖,成賢,菩薩,佛陀,老子,孔子,孟子,袁了凡,文天祥,呂祖,父母心,補運,謝恩,報親恩,答天恩,祭祀,掃墓,祖先,宗親,親戚,鄰里,鄉親,惜物,惜紙,惜穀,五穀,殊勝,契機,轉機,危機,生機,創新,效能,品質,效率,成果,績效,目標,願景,使命,精神,文化,核心,關鍵,重點,基礎,根本,活水,路標,燈塔,陽光,溫暖,希望,信心,勇氣,讀書會,講座,分享會,訓練,實習,經驗,傳承,接棒,火炬,光明,黑暗,無明,迷茫,絕望,珍惜,永恆,健康,藥石,舒壓,自省,覺知,能量,意識,共振,共榮,淨土,天堂,蓮花,法雨,滋潤,心田,播種,耕耘,灌溉,收成,豐收,果實,汗水,珍珠,傳世,聖壽,聖誕慶典,問心,志業,莊嚴,列聖,珍貴,寶藏,佛性,靈魂,善有善報,惡有惡報,天理昭彰,提醒,規勸,諄諄,婆心,仁慈,負責,專業,敬業,清醒,覺醒,禮貌,儀態,服裝,道衣,整潔,聖殿,內殿,拜亭,香爐,供桌,鮮花,素齋,叩首,跪拜,齋戒,慎言,慎行,慎思,慎獨,自律,自重,自愛,自強,自信,自足,自在,放下,捨得,煩憂,恐懼,驚慌,從容,淡定,平靜,柔軟,堅韌,剛毅,耐心,細心,用心,全心,學經,研經,聽經,持經,寫經,印經,傳經,演教,說法,宣教,弘法,化導,勸化,勸誡,警示,提點,啟發,啟迪,醒悟,改革,革新,改善,修正,調整,升級,完善,眾生,蒼生,萬物,靈性,性靈,慧命,法身,報身,化身,圓融,圓通,幹練,穩重,質樸,簡約,節儉,勤勞,勤快,勤勉,奮發,奮鬥,向上,增長,增進,增益,加強,鞏固,穩定,平衡,調和,調理,養生,保健,醫德,醫術,護理,康復,安頓,安穩,安心,安樂,安詳,安泰,敬香,捧茶,灑淨,辟邪,消災,解厄,賜福,延壽,誵經聲,肅穆,警戒線,善惡,功過,報應,功過格,五行,五臟';

const BUILT_IN_IGNORED = '也,的,和,在,而是,由於,首先,其次,以及,因為,所以,並非,雖然,不然,如果,然而,就是,之後,但是,那麼,此外,即使,而且,這個,那個,這些,那些,這次,那次,有,沒有,是,不是,可能,不可能,可以,不可以,應該,不應該';

// 特殊流水符號
const SPECIAL_SYMBOLS_STR = '㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩❶❷❸❹❺❻❼❽❾❿①②③④⑤⑥⑦⑧⑨⑩ⒶⒷⒸⒹⒺⒻⒼⒽⒾⒿ⒜⒝⒞⒟⒠⒡⒢⒣⒤⒥ⓐⓑⓒⓓⓔⓕⓖⓗⓘⓙ⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽';
const SPECIAL_SYMBOLS = new Set(Array.from(SPECIAL_SYMBOLS_STR));

// ── 工具函式 ──────────────────────────────────────────────
function parseList(str) {
  return str.split(/[,、，\n]/).map(s => s.trim()).filter(s => s.length > 0);
}

function getRandom(max) {
  return Math.random() < (max / 100);
}

// ── 主函式：標記文件 ────────────────────────────────────────
async function applyMarkup() {
  const btn = document.getElementById('btnApply');
  btn.disabled = true;
  btn.textContent = '⌛ 處理中…';
  setStatus('正在讀取文件內容…', 'info');

  try {
    await Word.run(async (context) => {
      const enhanceVal = parseInt(document.getElementById('enhanceDensity').value);
      const verbVal    = parseInt(document.getElementById('verbDensity').value);
      const shadingVal = parseInt(document.getElementById('shadingDensity').value);
      const bracketVal = parseInt(document.getElementById('bracketDensity').value);

      const extraImp = document.getElementById('extraImportant').value;
      const extraIgn = document.getElementById('extraIgnored').value;

      const finalIgnored   = new Set([...parseList(BUILT_IN_IGNORED), ...parseList(extraIgn)]);
      const finalImportant = [...parseList(BUILT_IN_IMPORTANT), ...parseList(extraImp)]
                               .sort((a, b) => b.length - a.length);

      // 取得文件 body
      const body = context.document.body;
      body.load('text');
      await context.sync();

      const fullText = body.text;
      if (!fullText.trim()) {
        setStatus('⚠️ 文件內容為空，請先輸入文字。', 'error');
        return;
      }

      // ── 先清除全文舊格式（底線、底色）──
      const bodyRange = body.getRange();
      bodyRange.font.underline = Word.UnderlineType.none;
      bodyRange.font.color     = '#000000';
      bodyRange.font.highlightColor = null;
      await context.sync();

      setStatus('正在分析詞語並套用標記…', 'info');

      // ── 搜尋並標記重要詞語 ──────────────────────────────
      let shadeTog = 0;

      for (const term of finalImportant) {
        if (!term || term.length < 1) continue;

        const searchResults = body.search(term, { matchCase: false, matchWholeWord: false });
        searchResults.load('items');
        await context.sync();

        for (const result of searchResults.items) {
          const useShade   = getRandom(shadingVal);
          const useBracket = getRandom(bracketVal);
          const is4Char    = term.length === 4;

          if (is4Char) {
            // 四字詞：前兩字 / 後兩字 交替底色
            applyHalfHighlight(result, shadeTog, shadingVal, context);
            shadeTog = 1 - shadeTog;
          } else {
            // 正常重要詞
            if (Math.random() < 0.5) {
              result.font.color     = '#cc0000';
              result.font.underline = Word.UnderlineType.thick;
              result.font.underlineColor = '#000000';
            } else {
              result.font.color     = '#000000';
              result.font.underline = Word.UnderlineType.thick;
              result.font.underlineColor = '#cc0000';
            }

            if (useShade) {
              result.font.highlightColor = shadeTog === 0 ? '#FFFF99' : '#BDD7EE';
              shadeTog = 1 - shadeTog;
            }

            if (useBracket) {
              // 在段落範圍插入括號，使用 insertText 前後包覆
              try {
                result.insertText('(', 'Before');
                result.insertText(')', 'After');
              } catch(e) { /* 某些情況無法插入，略過 */ }
            }
          }
        }
        await context.sync();
      }

      // ── 搜尋並標記次要詞語（verbDensity）────────────────
      // 用 Segmenter 從文字拆出詞，找到對應 range 並標記
      if (verbVal > 0) {
        await markSecondaryWords(body, context, finalImportant, finalIgnored, verbVal);
      }

      await context.sync();
      setStatus('✅ 標記完成！', 'ok');
    });

  } catch (err) {
    console.error(err);
    setStatus('❌ 發生錯誤：' + (err.message || err), 'error');
  }

  btn.disabled = false;
  btn.textContent = '✦ 標記當前文件';
}

// ── 四字詞輔助：嘗試對前後兩字分別標色（Word API 限制，近似做法）
function applyHalfHighlight(range, shadeTog, shadingVal, context) {
  // Word API 無法直接對 range 子字元操作底色（需用 paragraph 搜尋）
  // 此處以整體詞標記，交替底色以示區別
  const color1 = '#cc0000';
  const color2 = '#000000';
  if (Math.random() < 0.5) {
    range.font.color = color1;
  } else {
    range.font.color = color2;
  }
  range.font.underline = Word.UnderlineType.thick;
  if (Math.random() < (shadingVal / 100)) {
    range.font.highlightColor = shadeTog === 0 ? '#FFFF99' : '#BDD7EE';
  }
}

// ── 次要詞標記（Segmenter 拆詞後隨機套用灰線）────────────────
async function markSecondaryWords(body, context, importantList, ignoredSet, verbVal) {
  // 取得所有段落
  const paragraphs = body.paragraphs;
  paragraphs.load('items');
  await context.sync();

  const importantSet = new Set(importantList);

  for (const para of paragraphs.items) {
    para.load('text');
    await context.sync();

    const text = para.text;
    if (!text.trim()) continue;

    let segWords = [];
    try {
      const seg = new Intl.Segmenter('zh-TW', { granularity: 'word' });
      segWords = [...seg.segment(text)]
        .filter(s => s.isWordLike && s.segment.length >= 2 && s.segment.length <= 4);
    } catch(e) { continue; }

    for (const word of segWords) {
      const w = word.segment;
      if (ignoredSet.has(w) || importantSet.has(w)) continue;
      if (!getRandom(verbVal)) continue;

      // 在段落內搜尋該詞並套用灰底線紅字
      try {
        const results = para.search(w, { matchCase: true, matchWholeWord: false });
        results.load('items');
        await context.sync();

        for (const r of results.items) {
          r.font.color     = '#cc0000';
          r.font.underline = Word.UnderlineType.single;
          r.font.underlineColor = '#888888';
        }
        await context.sync();
      } catch(e) { /* 略過搜尋失敗的詞 */ }
    }
  }
}

// ── 清除所有標記 ───────────────────────────────────────────
async function clearMarkup() {
  setStatus('正在清除標記…', 'info');
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const range = body.getRange();
      range.font.underline      = Word.UnderlineType.none;
      range.font.color          = '#000000';
      range.font.highlightColor = null;
      await context.sync();
    });
    setStatus('✅ 已清除所有標記。', 'ok');
  } catch (err) {
    setStatus('❌ 清除失敗：' + (err.message || err), 'error');
  }
}
