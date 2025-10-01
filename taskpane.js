/* taskpane.js — Генератор официального текста из формального описания
   Профиль: профессиональный техписатель отчётов по пентесту.
   Работает с Ollama /api/chat (стрим), скрывает <think>... и блоки reasoning.
*/

(() => {
  const el = {
    host:  document.getElementById('host'),
    model: document.getElementById('model'),

    template: document.getElementById('template'),
    language: document.getElementById('language'),
    severity: document.getElementById('severity'),
    conciseness: document.getElementById('conciseness'),

    secContext: document.getElementById('secContext'),
    secObservation: document.getElementById('secObservation'),
    secTech: document.getElementById('secTech'),
    secImpact: document.getElementById('secImpact'),
    secEvidence: document.getElementById('secEvidence'),
    secRecom: document.getElementById('secRecom'),

    src:   document.getElementById('sourceText'),
    out:   document.getElementById('rewrittenText'),

    btnRewrite: document.getElementById('btnRewrite'),
    btnStop:    document.getElementById('btnStop'),
    btnClear:   document.getElementById('btnClear'),
    btnLoadSel: document.getElementById('btnLoadFromSelection'),
    btnInsert:  document.getElementById('btnInsertToDoc'),

    status:     document.getElementById('statusText'),
    statusDot:  document.getElementById('statusDot'),
  };

  let abortController = null;

  // ---- Init defaults & Office ready ----
  function initDefaults() {
    try {
      el.host.value  = localStorage.getItem('ollama_host')  || 'https://localhost:3000';
      el.model.value = localStorage.getItem('ollama_model') || 'deepseek-r1:14b';
      el.language.value = localStorage.getItem('writer_lang') || 'ru';
      el.template.value = localStorage.getItem('writer_tpl')  || 'finding';
      el.conciseness.value = localStorage.getItem('writer_conc') || 'concise';
      (localStorage.getItem('writer_sev') || '') && (el.severity.value = localStorage.getItem('writer_sev'));
      ['secContext','secObservation','secTech','secImpact','secEvidence','secRecom'].forEach(id=>{
        const v = localStorage.getItem('writer_'+id);
        if (v !== null) el[id].checked = v === '1';
      });
    } catch {}
  }
  initDefaults();

  [el.host, el.model, el.language, el.template, el.conciseness, el.severity].forEach(inp=>{
    inp.addEventListener('change',()=>{
      try {
        if (inp===el.host)  localStorage.setItem('ollama_host', el.host.value.trim());
        if (inp===el.model) localStorage.setItem('ollama_model', el.model.value.trim());
        if (inp===el.language) localStorage.setItem('writer_lang', el.language.value);
        if (inp===el.template) localStorage.setItem('writer_tpl', el.template.value);
        if (inp===el.conciseness) localStorage.setItem('writer_conc', el.conciseness.value);
        if (inp===el.severity) localStorage.setItem('writer_sev', el.severity.value);
      } catch {}
    });
  });
  ['secContext','secObservation','secTech','secImpact','secEvidence','secRecom'].forEach(id=>{
    el[id].addEventListener('change',()=>{
      try { localStorage.setItem('writer_'+id, el[id].checked ? '1' : '0'); } catch {}
    });
  });

  if (window.Office && Office.onReady) Office.onReady().catch(()=>{});

  // ---- UI helpers ----
  function setStatus(text, mode='ok'){
    el.status.textContent = text;
    el.statusDot.className = 'dot';
    if (mode==='running') el.statusDot.classList.add('running');
    else if (mode==='err') el.statusDot.classList.add('err');
    else el.statusDot.classList.add('ok');
  }

  function stripReasoning(s){
    if (!s) return s;
    return s
      .replace(/<think>[\s\S]*?<\/think>/gi,'')
      .replace(/```(?:thought|thinking|reasoning|xml)?[\s\S]*?```/gi,'')
      .replace(/\u0000/g,'')
      .trim();
  }

  // ---- Prompt builders ----
  function buildSystemPrompt(lang){
    const ru = [
      'Ты — профессиональный технический писатель по информационной безопасности.',
      'Пишешь разделы отчётов по пентесту в официально-деловом стиле.',
      'Требования:',
      '• Используй нейтральный тон, избегай разговорных оборотов.',
      '• Сохраняй точность терминов, чисел, артефактов, имён собственных.',
      '• Не выдумывай факты; не добавляй выводы, которых нет в исходных данных.',
      '• Структурируй текст по выбранному шаблону. Разделы выводи заголовками.',
      '• Если указана критичность, отрази её.',
      '• Никаких пояснений о ходе рассуждений. Выводи только итоговый текст.'
    ].join('\n');

    const en = [
      'You are a professional technical writer in information security.',
      'Produce penetration test report sections in a formal business style.',
      'Requirements:',
      '• Neutral, impersonal tone; avoid colloquialisms.',
      '• Preserve terminology, numbers, artifacts, and proper nouns.',
      '• Do not invent facts or add conclusions that are not provided.',
      '• Structure output using the selected template. Use section headings.',
      '• Reflect severity if provided.',
      '• Do not expose chain-of-thought. Output final text only.'
    ].join('\n');

    return lang==='en' ? en : ru;
  }

  function templateInstructions(tpl, lang){
    const RU = {
      finding: 'Шаблон: Finding/Уязвимость. Разделы: Контекст; Наблюдение; Технические детали; Влияние; Доказательства; Рекомендации.',
      exec:    'Шаблон: Executive Summary. Коротко и ясно: Контекст; Ключевые наблюдения; Риски/Влияние; Рекомендации в верхнеуровневом виде.',
      method:  'Шаблон: Методология/Шаги теста. Разделы: Цель; Окружение/Область; Выполненные шаги; Результаты; Ограничения; Вывод.'
    };
    const EN = {
      finding: 'Template: Finding/Vulnerability. Sections: Context; Observation; Technical Details; Impact; Evidence; Recommendations.',
      exec:    'Template: Executive Summary. Sections: Context; Key Observations; Risks/Impact; High-level Recommendations.',
      method:  'Template: Methodology/Test Steps. Sections: Objective; Scope/Environment; Steps Performed; Results; Limitations; Conclusion.'
    };
    const map = lang==='en'? EN: RU;
    return map[tpl] || map.finding;
  }

  function selectedSections(lang){
    const allRu = [
      el.secContext.checked && 'Контекст',
      el.secObservation.checked && 'Наблюдение',
      el.secTech.checked && 'Технические детали',
      el.secImpact.checked && 'Влияние',
      el.secEvidence.checked && 'Доказательства',
      el.secRecom.checked && 'Рекомендации',
    ].filter(Boolean);

    const allEn = [
      el.secContext.checked && 'Context',
      el.secObservation.checked && 'Observation',
      el.secTech.checked && 'Technical Details',
      el.secImpact.checked && 'Impact',
      el.secEvidence.checked && 'Evidence',
      el.secRecom.checked && 'Recommendations',
    ].filter(Boolean);

    return (lang==='en'? allEn: allRu).join('; ');
  }

  function concisenessHint(lang, level){
    const ru = {
      concise: 'Пиши кратко, без воды, 5–9 предложений на раздел.',
      balanced: 'Сбалансированная детализация, 8–14 предложений на раздел.',
      detailed: 'Подробно, но без лишней риторики; раскрывай технические нюансы.'
    };
    const en = {
      concise: 'Keep it concise, 5–9 sentences per section.',
      balanced: 'Balanced detail, 8–14 sentences per section.',
      detailed: 'Detailed yet focused; expand on technical nuances.'
    };
    return (lang==='en'? en: ru)[level || 'concise'];
  }

  function buildUserPrompt(formal, opts){
    const { lang, tpl, severity, conc } = opts;
    const secList = selectedSections(lang);
    const lines = [];

    if (lang==='en'){
      lines.push(
        'Transform the formal bullet points below into a polished, official business text for a penetration test report.',
        templateInstructions(tpl, lang),
        `Include only these sections: ${secList || 'Default for the template'}.`,
        concisenessHint(lang, conc),
        severity ? `Severity to reflect: ${severity}.` : 'If severity is absent, omit the section.',
        'Use headings for sections. Do not add any content not present in the formal notes.'
      );
      lines.push('\nFORMAL NOTES:\n' + formal);
      lines.push('\nOUTPUT: Provide final text only, no prefaces, no lists of requirements.');
    } else {
      lines.push(
        'Преобразуй приведённые ниже формальные пункты в выверенный официальный текст раздела отчёта по пентесту.',
        templateInstructions(tpl, lang),
        `Выводи только эти разделы: ${secList || 'Стандартные для шаблона'}.`,
        concisenessHint(lang, conc),
        severity ? `Отрази критичность: ${severity}.` : 'Если критичность не указана, опусти раздел.',
        'Разделы выводи заголовками. Не добавляй сведения, которых нет в исходных пунктах.'
      );
      lines.push('\nФОРМАЛЬНЫЕ ПУНКТЫ:\n' + formal);
      lines.push('\nВЫВОД: только финальный текст, без префиксов/пояснений.');
    }

    return lines.join('\n');
  }

  // ---- Streaming from Ollama ----
  async function* readChatStream(response){
    const reader = response.body.getReader();
    const decoder = new TextDecoder('utf-8');
    let buf = '';
    try{
      for(;;){
        const {value, done} = await reader.read();
        if (done) break;
        buf += decoder.decode(value, {stream:true});
        const lines = buf.split('\n'); buf = lines.pop() || '';
        for(const line of lines){
          const t = line.trim(); if(!t) continue;
          const jsonStr = t.startsWith('data:') ? t.slice(5).trim() : t;
          try { yield JSON.parse(jsonStr); } catch {}
        }
      }
      const tail = (buf||'').trim();
      if (tail){
        const jsonStr = tail.startsWith('data:') ? tail.slice(5).trim() : tail;
        try { yield JSON.parse(jsonStr); } catch {}
      }
    } finally { reader.releaseLock?.(); }
  }

  async function callOllamaChatStream({host, model, system, user}){
    const url = `${host.replace(/\/+$/,'')}/api/chat`;
    const body = {
      model,
      stream: true,
      messages: [
        { role:'system', content: system },
        { role:'user',   content: user }
      ],
      options: {
        temperature: 0.2
      }
    };
    abortController = new AbortController();
    const res = await fetch(url, {
      method:'POST',
      headers: { 'Content-Type':'application/json' },
      body: JSON.stringify(body),
      signal: abortController.signal
    });
    if (!res.ok || !res.body){
      const text = await res.text().catch(()=> '');
      throw new Error(`Ollama error: ${res.status} ${res.statusText}\n${text}`);
    }
    return res;
  }

  async function generate(){
    const formal = (el.src.value||'').trim();
    if (!formal){ setStatus('Введите формальные пункты/описание.', 'err'); return; }

    el.out.value = '';
    setStatus('Генерация...', 'running');
    el.btnRewrite.disabled = true;

    const lang = el.language.value;
    try{
      const res = await callOllamaChatStream({
        host: el.host.value.trim(),
        model: el.model.value.trim(),
        system: buildSystemPrompt(lang),
        user: buildUserPrompt(formal, {
          lang,
          tpl: el.template.value,
          severity: el.severity.value,
          conc: el.conciseness.value
        })
      });

      let acc = '';
      for await (const chunk of readChatStream(res)){
        if (chunk?.message?.content){
          acc += chunk.message.content;
          el.out.value = stripReasoning(acc);
          el.out.scrollTop = el.out.scrollHeight;
        }
      }
      el.out.value = stripReasoning(el.out.value);
      setStatus('Готово.', 'ok');
    } catch(err){
      console.error(err);
      setStatus(String(err.message||err), 'err');
    } finally {
      el.btnRewrite.disabled = false;
      abortController = null;
    }
  }

  function stopGen(){
    if (abortController){
      abortController.abort();
      abortController = null;
      el.btnRewrite.disabled = false;
      setStatus('Остановлено.', 'err');
    }
  }

  // ---- Word helpers ----
  async function loadSelectionFromWord(){
    if (!(window.Word && Word.run)){ setStatus('Word API недоступен.', 'err'); return; }
    try{
      await Word.run(async ctx=>{
        const sel = ctx.document.getSelection();
        sel.load('text'); await ctx.sync();
        if (sel.text && sel.text.trim()){
          el.src.value = sel.text;
          setStatus('Выделение загружено.', 'ok');
        } else setStatus('В документе ничего не выделено.', 'err');
      });
    } catch(e){
      console.error(e); setStatus('Не удалось получить выделение.', 'err');
    }
  }

  async function insertToWord(){
    const text = (el.out.value||'').trim();
    if (!text){ setStatus('Нет результата для вставки.', 'err'); return; }
    if (!(window.Word && Word.run)){ setStatus('Word API недоступен.', 'err'); return; }
    try{
      await Word.run(async ctx=>{
        const range = ctx.document.getSelection();
        range.insertText(text, Word.InsertLocation.replace);
        await ctx.sync();
      });
      setStatus('Вставлено в документ.', 'ok');
    } catch(e){
      console.error(e); setStatus('Не удалось вставить в документ.', 'err');
    }
  }

  // ---- Bindings ----
  el.btnRewrite.addEventListener('click', generate);
  el.btnStop.addEventListener('click', stopGen);
  el.btnClear.addEventListener('click', ()=>{ el.src.value=''; el.out.value=''; setStatus('Очищено.', 'ok'); });
  el.btnLoadSel.addEventListener('click', loadSelectionFromWord);
  el.btnInsert.addEventListener('click', insertToWord);
  el.src.addEventListener('keydown', e=>{
    if ((e.ctrlKey||e.metaKey) && e.key==='Enter'){ e.preventDefault(); generate(); }
  });
})();