/* Nonparticipant training is intentionally independent of the participant
 * workflow: no sign-in, D1 saves, event logging, or research submissions.
 * Only static curriculum and allowlisted media are fetched.
 */
(() => {
  'use strict';
  const KEY = 'aqg-open-training-v2-card';
  const MEDIA = 'https://dissertation-study-api.professor-jin.workers.dev/v1/training/media/2026-04-27-strict-content-a/videos/';
  const SWAP = {
    '05a-guided-practice-participant-v2.mp4': '05b-guided-practice-nonparticipant-v2.mp4',
    '06a-wrapping-up-participant-v2.mp4': '06b-wrapping-up-nonparticipant-v2.mp4'
  };
  const $ = id => document.getElementById(id);
  const create = (tag, text, className) => {
    const el = document.createElement(tag);
    if (text !== undefined) el.textContent = text;
    if (className) el.className = className;
    return el;
  };
  const put = (parent,child) => {parent.appendChild(child);return child;};
  const split = text => String(text||'').split(/\s*\|\|\s*/g).map(x=>x.trim()).filter(Boolean);
  const shuffle = values => {
    const a=[...values];
    for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[a[i],a[j]]=[a[j],a[i]];}
    return a;
  };
  let rows=[], idx=0, ended=false;
  function paragraphs(text,parent) {
    String(text||'').split(/\n\s*\n/g).forEach(part=>{
      if(!part.trim())return;
      const p=put(parent,create('p'));
      part.split(/(https?:\/\/[^\s]+)/g).forEach(piece=>{
        if (/^https?:\/\//.test(piece)) {
          const a=put(p,create('a',piece));a.href=piece;a.target='_blank';a.rel='noopener noreferrer';
        }else put(p,document.createTextNode(piece.replace(/\*\*/g,'')));
      });
    });
  }
  function adapted(row){
    const r={...row};
    if(r.header==='Welcome to this module'){
      r.body='This open module shows you how to use an AI chatbot and a question-writing app to create and refine nursing multiple-choice questions. No previous AI experience is required. The optional beginner guide explains AI chatbots and copying and pasting.\n\nThis version is not part of the study. There is no participant code, graded assessment, or research data submission. Your current place is remembered only in this browser.';
    }
    if(r.header==='Begin your practice session'){
      r.body='Open the nonparticipant practice app here:\n\nhttps://jin-dissertation.github.io/dissertation-app/np/npindex.html?mode=training\n\nKeep your AI chatbot open in another tab. Spend about 15 minutes exploring the app with a topic you teach. Practice goals (not graded):\n\n1. Try to accept at least 5 questions.\n2. Try each main question-refinement prompt at least once.\n3. Try each audit at least once.\n\nWhen you are finished, come back here to complete the module. No participant code is required.';
    }
    if(r.header==='Any Questions?'){
      r.body='You can find the Help and FAQ section in the nonparticipant question-writing app. For other questions you may contact tjin@crimson.ua.edu.';
    }
    return r;
  }
  function showText(r,parent){
    if(r.header)put(parent,create('h2',r.header));
    paragraphs(r.body,parent);
    if(String(r.media_url||'').startsWith('images/')){
      const img=put(parent,create('img'));
      img.src=r.media_url;img.alt=r.header||'Training illustration';img.loading='lazy';
    }
  }
  function showVideo(r,parent){
    put(parent,create('h2','Watch the video, then select Continue.'));
    const name=String(r.video_name||'').replace(/^videos\//,'');
    const selected=SWAP[name]||name;
    const player=put(parent,create('video'));player.controls=true;player.playsInline=true;player.preload='metadata';
    const source=put(player,create('source'));source.src=MEDIA+encodeURIComponent(selected);source.type='video/mp4';
    const track=put(player,create('track'));
    track.kind='subtitles';track.srclang='en';track.label='English';
    track.src=MEDIA+encodeURIComponent(selected.replace(/\.mp4$/i,'.vtt'));
    track.addEventListener('error',()=>track.remove());
    paragraphs('You can pause or replay the video whenever you like.',parent);
  }
  function showQuestion(r,parent){
    put(parent,create('h2','Knowledge check · ungraded'));
    put(parent,create('p',r.stem));
    const correct=r.type==='sba'?[r.correct_blob]:split(r.correct_blob);
    const options=shuffle([...correct,...split(r.distractors_blob)]);
    const checkGroup=put(parent,create('div'));
    const inputs=[];
    options.forEach((value,i)=>{
      const label=put(checkGroup,create('label',undefined,'choice'));
      const input=put(label,create('input'));input.type=r.type==='sba'?'radio':'checkbox';
      input.name='q'+idx; input.value=value;
      inputs.push(input);
      put(label,create('span',value));
    });
    const button=put(parent,create('button','Check answer','btn'));
    button.type='button';button.style.marginTop='10px';
    const feedback=put(parent,create('div',undefined,'answer'));feedback.hidden=true;
    button.addEventListener('click',()=>{
      const selected=inputs.filter(x=>x.checked).map(x=>x.value).sort();
      const match=JSON.stringify(selected)===JSON.stringify([...correct].sort());
      feedback.textContent=(match?'That matches the expected answer. ':'Expected answer: '+correct.join('; ')+'. ')+(r.feedback||'');
      feedback.hidden=false;button.disabled=true;inputs.forEach(x=>x.disabled=true);
    });
  }
  function showMatching(r,parent){
    put(parent,create('h2','Match the following · ungraded'));
    put(parent,create('p',r.stem));
    const pairs=String(r.correct_blob||'').split('\n').map(x=>x.split('::')).filter(x=>x.length===2);
    const pool=shuffle(pairs.map(x=>x[1])), selected=[];
    pairs.forEach(([term,answer])=>{
      const label=put(parent,create('label',undefined,'match'));
      put(label,create('span',term));
      const sel=put(label,create('select'));
      const ph=put(sel,create('option','Choose a description'));ph.value='';
      pool.forEach(value=>{const opt=put(sel,create('option',value));opt.value=value;});
      selected.push([sel,answer]);
    });
    const button=put(parent,create('button','Check matches','btn'));button.type='button';
    const feedback=put(parent,create('div',undefined,'answer'));feedback.hidden=true;
    button.addEventListener('click',()=>{
      const n=selected.filter(([input,answer])=>input.value===answer).length;
      feedback.textContent=n+' of '+pairs.length+' matched. '+(r.feedback||'');
      if(n!==pairs.length)feedback.textContent+=' Expected pairings: '+pairs.map(x=>x.join(' → ')).join('; ');
      feedback.hidden=false;button.disabled=true;selected.forEach(([input])=>input.disabled=true);
    });
  }
  function display(){
    if(!rows.length)return;
    ended=false;
    idx=Math.max(0,Math.min(idx,rows.length-1));
    localStorage.setItem(KEY,String(idx));
    const r=adapted(rows[idx]);
    $('counter').textContent='Card '+(idx+1)+' of '+rows.length;
    $('chapter').textContent=r.section_title||'Training';
    $('progress-fill').style.width=(100*(idx+1)/rows.length)+'%';
    const c=$('content');c.replaceChildren();
    if(r.type==='video')showVideo(r,c);
    else if(r.type==='sba'||r.type==='sata'){
      if(String(r.settings).includes('beginnerHelp')){
        put(c,AQGBeginnerHelp.renderCard({
          stem:r.stem,
          options:split(r.correct_blob),
          onContinue:next
        }));
      }else showQuestion(r,c);
    }else if(r.type==='matching'||r.type==='categorization')showMatching(r,c);
    else showText(r,c);
    $('back').disabled=idx===0;
    $('next').textContent=idx===rows.length-1?'Finish training':'Continue →';
    $('next').hidden=false;
    $('status').textContent='No grades or submitted answers';
    window.scrollTo(0,0);
  }
  function next(){
    if(idx<rows.length-1){idx++;display();return;}
    ended=true;
    $('content').replaceChildren();
    put($('content'),create('h2','Training complete'));
    paragraphs('You can now use the nonparticipant question-writing app whenever it is useful. Your responses here were not submitted to the study.', $('content'));
    const link=put($('content'),create('a','Open the question-writing app'));
    link.href='../np/npindex.html';link.className='btn';
    link.style.display='inline-block';link.style.textDecoration='none';
    $('next').hidden=true;$('status').textContent='Completed';
    localStorage.removeItem(KEY);
  }
  $('basics').addEventListener('click',()=>AQGBeginnerHelp.open());
  $('back').addEventListener('click',()=>{if(ended){display();return;}if(idx>0){idx--;display();}});
  $('next').addEventListener('click',next);
  fetch('module-content-v2.json',{cache:'no-store'})
    .then(response=>{if(!response.ok)throw Error('HTTP '+response.status);return response.json();})
    .then(data=>{
      if(!Array.isArray(data.rows)||data.rows.length<24)throw Error('Curriculum incomplete');
      rows=data.rows;
      const n=Number(localStorage.getItem(KEY));
      idx=Number.isInteger(n)&&n>=0&&n<rows.length?n:0;
      display();
    })
    .catch(error=>{
      const box=$('content');box.replaceChildren();
      put(box,create('p','Unable to load the training. '+error.message,'error'));
      $('next').disabled=true;
    });
})();
