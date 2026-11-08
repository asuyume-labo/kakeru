// ===== KAKERU わたしのいまここ（このファイル1つにまとめています） =====

/**
 * KAKERU わたしのいまここ — Webアプリ本体
 * 受検 → 採点 → スプレッドシートに記録 → 本人へPDFレポートをメール送付
 * 安全確認に該当した回答だけ、アスユメLabo.に通知します（対応はアスユメLabo.が判断）。
 */
const CONFIG = {
  SHEET_RESULTS: '受検結果',
  SHEET_PASSWORD: 'パスワード管理',
  REQUIRE_PASSWORD: true,              // 受検パスワード（月替わり）を使う場合は true
  ADMIN_EMAIL: 'kazumasa.okayasu@asuyume-labo.com', // 安全確認に該当したときの通知先
  REPLY_TO: 'kazumasa.okayasu@asuyume-labo.com',    // 受検者が返信したときの宛先
  SENDER_NAME: 'アスユメLabo. KAKERU',
  REPORT_FOLDER_ID: '',                // PDFの控えを保存するDriveフォルダID（空なら保存しない）
  RETENTION_YEARS: 3,                  // この年数を過ぎた記録とPDFを自動削除
  ID_PREFIX: 'KK',
  TZ: 'Asia/Tokyo'
};

const HEADERS = [
  'ID','受検日時','氏名','メール','所属','競技種目','年代','競技レベル','現在の状況','直近の大会','ケガによる練習制限',
  'タイプ','競技コンディション','心身のゆとり度',
  '①練習と回復のバランス','②練習の納得感','③課題の明確さ','④指導者との信頼','⑤試合での発揮度',
  '⑥競技満足度','⑦モチベーション','⑧レジリエンス','⑨心身のゆとり','⑩心の状態（フロー）',
  '自己効力感（推定）','自律性','有能感','関係性','バーンアウト兆候','スランプ兆候','動機づけ（外的優位）',
  '表示メッセージ','S1','S2','要確認','同意','同意日時','メール送信'
].concat(Array.from({length:40},(_,i)=>'No.'+String(i+1).padStart(2,'0'))).concat(['対応状況','対応メモ']);

/* ---------- 画面 ---------- */
function doGet(){
  const t = HtmlService.createTemplateFromFile('Index');
  t.dataJson = JSON.stringify(Object.assign(getClientData_(), {requirePassword: CONFIG.REQUIRE_PASSWORD}));
  t.logoSrc = 'data:image/png;base64,' + LOGO_B64;
  return t.evaluate()
    .setTitle('KAKERU わたしのいまここ')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

/* ---------- 外部サイトに置いた受検画面からの受付 ---------- */
function doPost(e){
  let res;
  try{
    const body = JSON.parse(e.postData.contents);
    if(body.action === 'checkPassword') res = {ok:true, valid:checkPassword(body.password)};
    else if(body.action === 'submit') res = submitResult(body.payload);
    else res = {ok:false, message:'不明なリクエストです。'};
  }catch(err){
    console.error(err);
    res = {ok:false, message:'処理中にエラーが起きました。時間をおいてもう一度お試しください。'};
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- パスワード ---------- */
function checkPassword(pw){
  if(!CONFIG.REQUIRE_PASSWORD) return true;
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_PASSWORD);
  if(!sh) return false;
  const ym = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM');
  const rows = sh.getDataRange().getValues().slice(1);
  return rows.some(([k,v]) => {
    const key = k instanceof Date ? Utilities.formatDate(k, CONFIG.TZ, 'yyyy/MM') : String(k).trim().replace(/-/g,'/');
    return key === ym && String(v).trim() !== '' && String(v).trim() === String(pw||'').trim();
  });
}

/* ---------- 受検の送信 ---------- */
function submitResult(p){
  // 入力チェック
  const a = {}, attrs = p && p.attrs || {};
  for(let n=1;n<=40;n++){
    const v = Number(p.answers && p.answers[n]);
    if(!(v>=1 && v<=5 && Number.isInteger(v))) return {ok:false, message:'未回答の質問があります。回答に戻って確認してください。'};
    a[n] = v;
  }
  const sa = {S1: p.safety && p.safety.S1, S2: p.safety && p.safety.S2};
  if(!['yes','no'].includes(sa.S1) || !['yes','no'].includes(sa.S2)) return {ok:false, message:'最後の2つの質問に回答してください。'};
  const email = String(attrs.email||'').trim();
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return {ok:false, message:'メールアドレスの形式を確認してください。'};
  if(!String(attrs.name||'').trim()) return {ok:false, message:'お名前を入力してください。'};
  if(!checkPassword(p.password)) return {ok:false, message:'受検パスワードの有効期限が切れています。最初の画面からやり直してください。'};

  // 採点
  const r = computeResult_(a, sa);
  const now = new Date();
  r.id = CONFIG.ID_PREFIX + '-' + Utilities.formatDate(now, CONFIG.TZ, 'yyyyMMdd') + '-' + randomCode_(6);
  r.date = Utilities.formatDate(now, CONFIG.TZ, 'yyyy年M月d日');
  const needsCheck = r.safety ? 'S1該当' : (r.restFlag ? '休息フラグ' : '');

  // 記録（メール送信より先に保存する）
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let rowNo;
  try{
    const sh = ensureSheet_();
    const consentAt = p.consentAt ? new Date(p.consentAt) : now;
    const row = [
      r.id, now, attrs.name, email, attrs.team||'', attrs.sport||'', attrs.age||'', attrs.level||'', attrs.season||'', attrs.days||'', attrs.injury||'',
      r.type, r.condition, r.yutori
    ].concat([1,2,3,4,5,6,7,8,9,10].map(c=>r.s[c])).concat([
      r.efficacy, r.sdt['自律性'], r.sdt['有能感'], r.sdt['関係性'], r.burnout, r.slump?'あり':'なし', r.motivTilt?'該当':'',
      r.chosen.map(c=>c.id).join(' '), sa.S1==='yes'?'はい':'いいえ', sa.S2==='yes'?'はい':'いいえ', needsCheck, '同意済', consentAt, '送信中'
    ]).concat(Array.from({length:40},(_,i)=>a[i+1])).concat(['','']);
    sh.appendRow(row);
    rowNo = sh.getLastRow();
    sh.getRange(rowNo, 2).setNumberFormat('yyyy/MM/dd HH:mm:ss');
    sh.getRange(rowNo, HEADERS.indexOf('同意日時')+1).setNumberFormat('yyyy/MM/dd HH:mm:ss');
    if(needsCheck) sh.getRange(rowNo, HEADERS.indexOf('要確認')+1).setBackground('#f8e3e3');
  } finally { lock.releaseLock(); }

  // 本人へレポート送付
  let mailed = false;
  try{
    const logo = Utilities.newBlob(Utilities.base64Decode(LOGO_B64), 'image/png', 'logo.png');
    const pdf = Utilities.newBlob(buildReportHtml_(r, attrs, a, sa, 'data:image/png;base64,'+LOGO_B64, true), 'text/html', 'report.html')
      .getAs('application/pdf').setName('KAKERU_フィードバックシート_' + r.id + '.pdf');
    MailApp.sendEmail({
      to: email,
      subject: '【KAKERU】あなたのフィードバックシート',
      htmlBody: buildMailHtml_(r, attrs) ,
      attachments: [pdf],
      inlineImages: {logo: logo},
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.REPLY_TO
    });
    mailed = true;
    if(CONFIG.REPORT_FOLDER_ID) DriveApp.getFolderById(CONFIG.REPORT_FOLDER_ID).createFile(pdf);
  }catch(e){
    console.error('mail failed', e);
  }
  SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_RESULTS)
    .getRange(rowNo, HEADERS.indexOf('メール送信')+1).setValue(mailed ? '送信済' : '失敗');

  // 安全確認に該当したときだけアスユメLabo.へ通知（回答の中身は書かない）
  if(needsCheck){
    try{
      MailApp.sendEmail({
        to: CONFIG.ADMIN_EMAIL,
        subject: '【KAKERU】確認が必要な回答があります（' + r.id + '）',
        body: [
          'KAKERUで確認が必要な回答がありました。内容を確認し、対応を判断してください。',
          '', '受検ID：' + r.id, '受検日時：' + Utilities.formatDate(now, CONFIG.TZ, 'yyyy/MM/dd HH:mm'),
          '区分：' + needsCheck + (needsCheck==='S1該当' ? '（安全確認S1に「はい」）' : '（⑨心身のゆとりが20点以下かつ「充電が必要」が5カテゴリ以上）'),
          '', '記録シート：' + SpreadsheetApp.getActive().getUrl()
        ].join('\n'),
        name: CONFIG.SENDER_NAME
      });
    }catch(e){ console.error('notify failed', e); }
  }

  return {ok:true, result:r, mailed:mailed};
}

function randomCode_(n){
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; let s = '';
  for(let i=0;i<n;i++) s += c.charAt(Math.floor(Math.random()*c.length));
  return s;
}

/* ---------- 初期設定（最初に1回だけ実行） ---------- */
function setup(){
  ensureSheet_();
  const ss = SpreadsheetApp.getActive();
  let pw = ss.getSheetByName(CONFIG.SHEET_PASSWORD);
  if(!pw){
    pw = ss.insertSheet(CONFIG.SHEET_PASSWORD);
    pw.getRange(1,1,1,2).setValues([['年月','パスワード']]).setFontWeight('bold');
    pw.getRange('A:A').setNumberFormat('@');
    const rows = [], d = new Date();
    for(let i=0;i<12;i++){
      const m = new Date(d.getFullYear(), d.getMonth()+i, 1);
      rows.push([Utilities.formatDate(m, CONFIG.TZ, 'yyyy/MM'), String(1000+Math.floor(Math.random()*9000))]);
    }
    pw.getRange(2,1,rows.length,2).setValues(rows);
  }
  // 古い記録の自動削除（毎日3時台）
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='purgeOldRecords').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('purgeOldRecords').timeBased().everyDays(1).atHour(3).create();
}

function ensureSheet_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SHEET_RESULTS);
  if(!sh){
    sh = ss.insertSheet(CONFIG.SHEET_RESULTS, 0);
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#14212b').setFontColor('#ffffff');
    sh.setFrozenRows(1); sh.setFrozenColumns(3);
  }
  return sh;
}

/* ---------- 保存期間を過ぎた記録の削除 ---------- */
function purgeOldRecords(){
  const limit = new Date(); limit.setFullYear(limit.getFullYear() - CONFIG.RETENTION_YEARS);
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_RESULTS);
  if(sh && sh.getLastRow() > 1){
    const dates = sh.getRange(2, 2, sh.getLastRow()-1, 1).getValues();
    for(let i=dates.length-1;i>=0;i--){
      const d = dates[i][0];
      if(d instanceof Date && d < limit) sh.deleteRow(i+2);
    }
  }
  if(CONFIG.REPORT_FOLDER_ID){
    const it = DriveApp.getFolderById(CONFIG.REPORT_FOLDER_ID).getFiles();
    while(it.hasNext()){ const f = it.next(); if(f.getDateCreated() < limit) f.setTrashed(true); }
  }
}


/**
 * KAKERU わたしのいまここ — 設問・採点・文例（サーバーとブラウザで共通の定義）
 * 文言や閾値を変えるときは、このファイルだけを直せば画面・メール・シートすべてに反映されます。
 */
const CATS = {
 1:{name:'練習と回復のバランス'},2:{name:'練習の納得感'},3:{name:'課題の明確さ'},4:{name:'指導者との信頼'},
 5:{name:'試合での発揮度'},6:{name:'競技満足度'},7:{name:'モチベーション'},8:{name:'レジリエンス'},
 9:{name:'心身のゆとり'},10:{name:'心の状態（フロー）'}
};
const MARK = ['','①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩'];
const ITEMS = [
 [1,1,'練習の量や強度は、今の自分の体力に見合っていると感じる',0],
 [2,1,'練習の疲れが翌日まで残ることが多い',1],
 [3,1,'休養や睡眠を十分にとれている',0],
 [4,1,'練習がきつすぎて、こなすだけで精一杯になっている',1],
 [5,2,'今取り組んでいる練習の目的を理解している',0],
 [6,2,'練習内容について、自分の意見を伝えたり選んだりできる',0],
 [7,2,'「やらされている」と感じながら練習することが多い',1],
 [8,2,'今の練習が自分の成長につながっている実感がある',0],
 [9,3,'自分が今取り組むべき課題がはっきりしている',0],
 [10,3,'課題を克服するための具体的な方法がわかっている',0],
 [11,3,'何をどう直せばよいのかわからず迷うことがある',1],
 [12,3,'課題が多すぎて、どこから手をつければよいかわからない',1],
 [13,4,'コーチ・監督を信頼している',0],
 [14,4,'コーチ・監督は、自分のことを理解しようとしてくれている',0],
 [15,4,'コーチ・監督と、目指している方向が一致している',0],
 [16,4,'コーチ・監督に、本音や不安を伝えることができる',0],
 [17,5,'最近の試合や記録で、自分の力を発揮できている',0],
 [18,5,'練習でできていることが、試合では出せないと感じる',1],
 [19,5,'試合前の不安や緊張で、調子を崩すことがある',1],
 [20,5,'結果が出ないことへの焦りが強い',1],
 [21,6,'競技を続けていることに満足している',0],
 [22,6,'競技をしている時間を楽しめている',0],
 [23,6,'競技と、学業・仕事・私生活とのバランスがとれている',0],
 [24,6,'競技への興味や関心が、以前より薄れている',1],
 [25,7,'この競技そのものが好きだから取り組んでいる',0],
 [26,7,'自分で決めた目標に向かって取り組めている',0],
 [27,7,'周りの期待に応えるため、または怒られないために続けている',1],
 [28,7,'なぜこの競技を続けているのか、わからなくなることがある',1],
 [29,8,'失敗やミスの後でも、気持ちを切り替えられる',0],
 [30,8,'うまくいかない時期があっても、また取り戻せると思える',0],
 [31,8,'不安や悔しさがあっても、やるべきプレーに集中できる',0],
 [32,8,'困ったときに相談できる人がいる',0],
 [33,9,'気分が落ち込んだり、イライラしたりすることが多い',1],
 [34,9,'寝つけない、夜中に目が覚めるなど、睡眠の不調がある',1],
 [35,9,'競技のことが頭から離れず、休んでいても気が休まらない',1],
 [36,9,'心も体も疲れ切っていると感じる',1],
 [37,10,'試合や練習中、結果やミスにとらわれず、目の前のプレーに集中できている',0],
 [38,10,'周りの状況や人からの評価に、気持ちが揺さぶられやすい',1],
 [39,10,'自分の機嫌を、自分で整える方法を持っている',0],
 [40,10,'過去の失敗や先の結果を考えすぎて、「今」に集中できないことがある',1]
].map(([no,cat,text,rev])=>({no,cat,text,rev:!!rev}));
const ITEM = Object.fromEntries(ITEMS.map(i=>[i.no,i]));
// 画面表示順：カテゴリが混ざるように各カテゴリの1問目→2問目…の順
const ORDER = []; for(let r=0;r<4;r++) for(let c=1;c<=10;c++) ORDER.push((c-1)*4+r+1);
const LABELS = ['まったく\nあてはまらない','あまり\nあてはまらない','どちらとも\nいえない','やや\nあてはまる','とても\nあてはまる'];
const SAFETY = [
 {id:'S1',text:'最近2週間で、消えてしまいたい・自分を傷つけたいと思ったことがある'},
 {id:'S2',text:'指導者やチーム内で、暴言・威圧・体罰など、怖い・つらいと感じる関わりがある'}
];
const COMMENT = {
 1:['練習と回復のバランスが上手にとれています。今の休み方があなたの武器です','体がもう少し回復の時間を求めています。睡眠と休養日を意識すると、練習の質がさらに上がります','いまは回復を優先するタイミングです。指導者やトレーナーと練習量を相談すると、次の伸びにつながります'],
 2:['練習の目的を理解し、主体的に取り組めています','目的がはっきりすると、同じ練習でも手応えが大きくなります。ねらいを一つ確認してみましょう','自分で選べる部分を一つつくると、練習がぐっと自分のものになります'],
 3:['課題と方法がクリアです。この見通しが成長を後押ししています','課題を「今月の一つ」に絞ると、前進がはっきり感じられます','課題を誰かと一緒に整理すると、進む道が見えてきます。相談すること自体が前進です'],
 4:['指導者との信頼関係が、あなたの大きな土台になっています','聞きたいことを一つ伝えると、指導者との方向性がさらにそろいます','話しやすい第三者に先に気持ちを整理してもらうと、伝え方が見えてきます'],
 5:['試合で力を発揮できています。うまくいった時の感覚を言葉にしておきましょう','練習で培った力は確かにあります。試合前のルーティンを整えると、もっと出しやすくなります','力をため込んでいる時期です。過去にうまくいった試合の状態を振り返ると、出し方を思い出せます'],
 6:['競技を楽しみ、生活とのバランスもとれています','競技以外の時間も大切にすると、競技の楽しさがもう一段増します','少し距離をとって心を休めることも、競技を長く楽しむための大切な時間です'],
 7:['自分の意思で、前向きに取り組めています','「始めた理由」を思い出すと、自分のための力がもっと湧いてきます','いまは動機を見つめ直すチャンスです。「なぜ続けてきたのか」を一緒に考えてみましょう'],
 8:['切り替えと立ち直りの力が、あなたの強みです','切り替えの合図を一つ決めると、さらに素早く立て直せます','一人で抱えず、相談できる人に話すことで、立て直す力が戻ってきます'],
 9:['心と体にゆとりがあり、安定しています','少しお疲れ気味です。睡眠と休養を優先すると、ゆとりが戻ってきます','心と体がしっかり休むことを求めています。専門家（医療機関・心理職）に話すと、回復の近道になります'],
 10:['揺らがず・とらわれず、今のプレーに集中できています。その状態をつくる自分なりの方法を大切にしましょう','プレーの前に「今・ここ」に戻る合図を一つ決めると、心がさらに安定します','自分の機嫌を整える小さな習慣を一つ持つと、心がぐっと軽くなります']
};
const ONEPOINT = {
 1:['疲れに気づけていること自体が、体を守る力です。今夜は「よく休めた」と言える夜を一つ作りましょう','生理的・情動的状態'],
 2:['今日の練習の目的を、自分の言葉で一行書いてみましょう。意味がわかると同じ練習でも手応えが変わります','達成体験'],
 3:['この1年でできるようになったことを3つ書き出してみましょう。課題が見えるのは、成長してきた人だけです','達成体験'],
 4:['信頼は一度の会話では変わりません。「一つだけ質問する」ことから始めれば十分です','言語的説得'],
 5:['いちばん力を出せた試合の「体の感覚」を思い出してください。その力はなくなっていません','達成体験（ピーク体験）'],
 6:['競技を楽しいと感じた瞬間を一つ思い出してみましょう。そこがあなたの原点です','達成体験'],
 7:['「なぜまだやめていないのか」に答えてみてください。そこに残っている理由が、あなたの本当の動機です','言語的説得（自己対話）'],
 8:['悔しさを感じるのは、まだ本気だからです。その気持ちは次の一歩の燃料になります','生理的・情動的状態（読み替え）'],
 9:['自分の状態に気づけたことが、もう最初の一歩です。休むことも、強くなるための練習の一部です','生理的・情動的状態'],
 10:['結果は選べなくても、心の状態は自分で選べます。今日、機嫌よくプレーできた瞬間を一つ思い出してみましょう','生理的・情動的状態（自己調整）']
};
const SELFCARE = {
 1:'寝る時間を30分早め、週に1日は完全休養日をつくる',
 2:'練習前に「今日のねらい」を一つだけ自分で決めてから入る',
 3:'課題を紙に書き出し、今月取り組むものに一つだけ丸をつける',
 4:'次の練習後、コーチ・監督に聞きたいことを一つだけ質問する',
 5:'うまくいった試合の前日・当日の過ごし方を書き出し、ルーティンに入れる',
 6:'競技と関係のない、自分が楽しめる時間を週に一度つくる',
 7:'競技を始めたころの「楽しかった場面」を一つ思い出して書く',
 8:'ミスの後に使う切り替えの合図（深呼吸1回、手を叩くなど）を決める',
 9:'寝る前1時間は競技の情報から離れ、睡眠を最優先にする',
 10:'プレー前に「今・ここ」に戻るための一言か動作を一つ決める'
};
const TYPES = {
 '充実型':{cls:'good',line:'環境にも恵まれ、自分の中のエネルギーも豊かな好循環の状態です。',msg:'今の好循環は、あなたの積み重ねの結果です。うまくいっている理由を言葉にして、次のステージに備えましょう'},
 '挑戦型':{cls:'warn',line:'厳しい環境の中でも、自分の中のエネルギーで力強く挑めている状態です。',msg:'厳しい状況に挑み続けている力は本物です。その力を長く使うために、休む勇気も持ちましょう'},
 '充電型':{cls:'warn',line:'環境は整っています。いまは、自分の中のエネルギーを蓄え直す時期です。',msg:'充電の時期は、次に伸びるための準備期間です。小さな「できた」を一つずつ積み重ねていきましょう'},
 '回復優先型':{cls:'crit',line:'いまは心と体を休め、回復を最優先にすることで、また力が戻ってくる時期です。',msg:'ここまで頑張ってきた自分を、まず認めてあげてください。相談することは、前に進むための力です'}
};
// 回答パターン（グループ：esteem=自己肯定感系 / eff=自己効力感・モチベーション系）
const PATTERNS = [
 {id:'P11',g:'esteem',base:'自己肯定感・自己効力感',cond:c=>c.s[10]>=70&&c.s[5]<50,ev:[37,17],
  msg:'結果がまだついてこなくても、揺らがずに今のプレーに向き合えています。それは結果より先に整えるべき土台です'},
 {id:'P3',g:'esteem',base:'自己肯定感',cond:c=>c.a[20]>=4&&c.a[26]>=4,ev:[26,20],
  msg:'焦るのは、自分の目標を本気で追いかけているからです。その熱量はそのままに、目の前の一つに絞ってみましょう'},
 {id:'P4',g:'esteem',base:'自己肯定感',cond:c=>c.s[1]<50&&c.s[2]>=70,ev:[5,4],
  msg:'練習の意味を理解して全力で向き合えるのは大きな力です。その力を長く使うために、休むことも練習に入れましょう'},
 {id:'P7',g:'esteem',base:'自己肯定感→自律性',cond:c=>c.a[27]>=4&&c.a[28]<=2,ev:[27],
  msg:'周りの期待に応えようとする責任感は、あなたの誠実さです。その上で「自分のための目標」も一つ持ってみませんか'},
 {id:'P1',g:'eff',base:'自己効力感',cond:c=>c.s[5]<50&&c.s[8]>=70,ev:[29,30],
  msg:'結果が出ない時期でも気持ちを立て直せている。それは、結果より先に育っているあなたの本当の力です'},
 {id:'P12',g:'eff',base:'自己効力感（持続性）',cond:c=>c.s[5]>=70&&c.s[10]<50,ev:[17,40],
  msg:'結果を出している力は本物です。その力を長く発揮するために、結果だけでなく心の状態にも目を向けてみましょう'},
 {id:'P5',g:'eff',base:'自己効力感',cond:c=>c.a[9]>=4&&c.a[10]<=2,ev:[9],
  msg:'何を直すべきかが見えているのは、もう半分進んでいる証拠です。方法は人に聞いて増やせます'},
 {id:'P2',g:'eff',base:'モチベーション（自律性）',cond:c=>c.a[25]>=4&&(c.s[5]<50||c.s[6]<50),ev:[25],
  msg:'苦しい時期にも「好き」が残っているのは、競技を続ける一番強い理由を持っている証拠です'},
 {id:'P9',g:'eff',base:'関係性',cond:c=>mean([13,14,15,16].map(n=>c.a[n]))>=4,ev:[13,16],
  msg:'指導者と信頼で結ばれていることは、苦しい時期を越えるための大きな支えです'},
 {id:'P8',g:'eff',base:'関係性',cond:c=>c.a[32]>=4,ev:[32],
  msg:'あなたには相談できる人がいます。その人に、今日のチェックで気づいたことを一つ話してみましょう'},
 {id:'P6',g:'eff',base:'自己効力感（達成体験）',cond:c=>c.a[29]>=4||c.a[30]>=4,ev:[29,30],
  msg:'これまでも立て直してきた経験があります。その経験は、次の壁でも必ず使えます'},
 {id:'P10',g:'esteem',base:'自己肯定感',cond:c=>c.careCount>=5,ev:[],
  msg:'今の自分の状態に正直に向き合って回答したこと自体が、回復への第一歩です'}
];


function mean(xs){return xs.reduce((p,q)=>p+q,0)/xs.length}
const toScore = m => Math.round((m-1)/4*100);
function band(v){return v>=70?['好調','good',0]:v>=50?['もう一歩','warn',1]:['充電が必要','crit',2]}
function computeResult_(a, sa){
  const good = n => ITEM[n].rev ? 6-a[n] : a[n];
  const s = {};
  for(let c=1;c<=10;c++) s[c]=toScore(mean(ITEMS.filter(i=>i.cat===c).map(i=>good(i.no))));
  const yutori = s[9];
  const condition = Math.round(mean([1,2,3,4,5,6,7,8,10].map(c=>s[c])));
  const envMean = Math.round(mean([1,2,3,4,5,6].map(c=>s[c])));
  const resMean = Math.round(mean([7,8,10].map(c=>s[c])));
  const envHigh = envMean>=60, resHigh = resMean>=60;
  const type = resHigh ? (envHigh?'充実型':'挑戦型') : (envHigh?'充電型':'回復優先型');
  const sub = ns => toScore(mean(ns.map(good)));
  const bo = {消耗感:sub([2,4,36]),達成感:sub([8,17]),競技価値:sub([24,28])};
  const boLow = Object.values(bo).filter(v=>v<=49).length;
  const burnout = boLow===3?'フラグ':boLow===2?'注意':'なし';
  const slump = s[5]<=49 && [3,7,8,10].some(c=>s[c]<=49);
  const auto = mean([a[25],a[26]]), ctrl = mean([a[27],a[28]]);
  const motivTilt = ctrl>auto;
  const careCount = Object.values(s).filter(v=>v<50).length;
  const efficacy = sub([8,10,17,26,29,30,31]);
  const sdt = {自律性:sub([6,7,26]),有能感:sub([8,10,17]),関係性:sub([13,16,32])};
  const safety = sa.S1==='yes';
  const restFlag = !safety && s[9]<=20 && careCount>=5;
  const flags = [];
  if(sa.S1==='yes') flags.push([1,'S1「はい」','相談先を冒頭に表示／アスユメLabo.へ通知（対応はアスユメが判断）']);
  if(restFlag) flags.push([2,'⑨心身のゆとり '+s[9]+'点（20点以下）かつ「充電が必要」'+careCount+'カテゴリ','休息のすすめを表示／アスユメLabo.へ通知（対応はアスユメが判断）']);
  if(sa.S2==='yes') flags.push([3,'S2「はい」','第三者の相談窓口を表示（団体へは報告しない）']);
  if(burnout==='フラグ') flags.push([4,'バーンアウト兆候フラグ','休養と負荷の見直しを提案／面談を案内']);
  const ctx = {a,s,careCount};
  const matched = PATTERNS.filter(p=>p.cond(ctx)).map(p=>p.id);
  let chosen = [];
  if(safety){ chosen = [PATTERNS.find(p=>p.id==='P10')]; }
  else{
    const e1 = PATTERNS.find(p=>p.g==='eff'&&matched.includes(p.id));
    const e2 = PATTERNS.find(p=>p.g==='esteem'&&p.id!=='P10'&&matched.includes(p.id));
    chosen = [e1,e2].filter(Boolean);
    if(chosen.length<2){
      const extra = PATTERNS.find(p=>matched.includes(p.id)&&!chosen.includes(p)&&!(p.id==='P10'&&chosen.length));
      if(extra) chosen.push(extra);
    }
  }
  const sorted = Object.entries(s).map(([c,v])=>[+c,v]).sort((x,y)=>y[1]-x[1]);
  const strengths = sorted.slice(0,2), growth = sorted.slice(-2).reverse();
  chosen = chosen.map(p=>({id:p.id,g:p.g,base:p.base,msg:p.msg,ev:p.ev}));
  return {s,restFlag,yutori,condition,envMean,resMean,type,bo,burnout,slump,motivTilt,auto,ctrl,careCount,efficacy,sdt,safety,flags,matched,chosen,strengths,growth};
}


/** ブラウザに渡すデータ（採点ロジックは含めない） */
function getClientData_(){
  return {CATS, MARK, ITEMS, ORDER, LABELS, SAFETY, COMMENT, ONEPOINT, SELFCARE, TYPES};
}


/**
 * KAKERU — メール本文とPDFレポート（フィードバックシート）
 * PDFはHTMLから変換するため、表とインラインのスタイルだけで組んでいます。
 */
const RC = {ink:'#14212b', muted:'#566775', line:'#d6dee4', soft:'#e8eef2', accent:'#2453a6', accentSoft:'#e3ebf8',
  good:'#2b8653', goodSoft:'#e2f2e8', warn:'#a77406', warnSoft:'#f8efda', crit:'#5b5bb0', critSoft:'#e9e9f7', band:'#14212b'};
const FONT = "'Hiragino Sans','Hiragino Kaku Gothic ProN','Noto Sans JP','Yu Gothic','Meiryo',sans-serif";
const ANSWER_WORDS = ['まったくあてはまらない','あまりあてはまらない','どちらともいえない','ややあてはまる','とてもあてはまる'];
const TEL = 'いのちの電話 0570-783-556 ／ こころの健康相談統一ダイヤル 0570-064-556';

function esc_(t){return String(t==null?'':t).replace(/[&<>"]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[ch]))}
function pill_(label, key){return `<span style="display:inline-block;padding:1px 9px;border-radius:9px;font-size:11px;font-weight:bold;background:${RC[key+'Soft']};color:${RC[key]}">${esc_(label)}</span>`}
function h2_(t){return `<div style="font-size:15px;font-weight:bold;color:${RC.ink};margin:22px 0 8px;padding-bottom:4px;border-bottom:1px solid ${RC.line}">${esc_(t)}</div>`}
function bar_(v, key){
  const w = Math.max(v, 2);
  return `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse"><tr>
    <td style="width:${w}%;height:9px;background:${RC[key]};font-size:1px">&nbsp;</td>
    <td style="width:${100-w}%;height:9px;background:${RC.soft};font-size:1px">&nbsp;</td></tr></table>`;
}
function header_(logoSrc, r){
  return `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:${RC.band}"><tr>
    <td style="padding:12px 16px"><img src="${logoSrc}" alt="ASUYUME LABO." style="height:18px;display:block"></td>
    <td style="padding:12px 16px;text-align:right;color:#a9b7c2;font-size:11px">${r?('受検ID '+esc_(r.id)+'　'+esc_(r.date)):''}</td></tr></table>`;
}
function evText_(p, a){
  if(!p.ev || !p.ev.length) return 'あなたは、すべての質問に正直に向き合って回答しています。';
  return 'あなたは' + p.ev.map(n=>`「${ITEM[n].text}」に${ANSWER_WORDS[a[n]-1]}`).join('、') + 'と答えています。';
}

/** PDF（フィードバックシート全体） */
function buildReportHtml_(r, attrs, a, sa, logoSrc){
  const T = TYPES[r.type];
  const yB = band(r.yutori), cB = band(r.condition);
  const sdtMin = Object.keys(r.sdt).sort((x,y)=>r.sdt[x]-r.sdt[y])[0];
  const opCats = r.safety ? [9] : r.growth.map(g=>g[0]);
  const careCats = r.safety ? [9] : Object.keys(r.s).map(c=>[+c, r.s[c]]).filter(x=>x[1]<70).sort((x,y)=>x[1]-y[1]).slice(0,3).map(x=>x[0]);
  const careList = (careCats.length ? careCats : [r.strengths[0][0]]).map(c=>SELFCARE[c]);
  const attrLine = [attrs.sport, attrs.age, attrs.level, attrs.season].filter(Boolean).join('／');
  const effWord = r.efficacy>=70?'しっかりある':r.efficacy>=50?'育ちつつある':'これから育てていく時期';
  const p = s => `<p style="margin:4px 0;font-size:12px;line-height:1.75;color:${RC.ink}">${s}</p>`;
  const small = s => `<p style="margin:4px 0;font-size:11px;line-height:1.7;color:${RC.muted}">${s}</p>`;
  let h = `<html><head><meta charset="utf-8"></head><body style="margin:0;font-family:${FONT};color:${RC.ink}">`;
  h += header_(logoSrc, r);
  h += `<div style="padding:16px 22px">`;
  h += `<div style="font-size:11px;letter-spacing:2px;color:${RC.muted}">KAKERU FEEDBACK SHEET</div>
        <div style="font-size:22px;font-weight:bold;margin:2px 0">${esc_(attrs.name)} さんの競技コンディション</div>
        ${attrLine?small(esc_(attrLine)):''}`;

  if(r.safety){
    h += `<div style="margin:12px 0;padding:12px 14px;border:1px solid ${RC.line};border-left:3px solid ${RC.accent}">
      <div style="font-size:13px;font-weight:bold">よかったら、誰かに話してみませんか</div>
      ${p('正直に答えてくれて、ありがとうございます。気持ちを言葉にするだけで、少し楽になることがあります。家族や友人、チームドクターなど、話しやすい人に声をかけてみてください。')}
      ${small('電話で相談できる窓口：'+TEL)}</div>`;
  } else if(r.restFlag){
    h += `<div style="margin:12px 0;padding:12px 14px;border:1px solid ${RC.line}">
      <div style="font-size:13px;font-weight:bold">いまは、しっかり休むことがいちばんの練習です</div>
      ${p('心と体が休息を求めているようです。睡眠と休養を優先し、疲れが続くときは身近な人や専門家に話してみてください。')}</div>`;
  }

  // サマリー
  h += `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin-top:10px;border:1px solid ${RC.line}"><tr>
    <td style="padding:12px 14px;width:46%;vertical-align:top;border-right:1px solid ${RC.line}">
      <div style="font-size:11px;color:${RC.muted}">あなたのタイプ</div>
      <div style="font-size:24px;font-weight:bold;margin:2px 0">${esc_(r.type)}</div>
      <div style="font-size:11px;color:${RC.muted};line-height:1.6">${esc_(T.line)}</div></td>
    <td style="padding:12px 14px;vertical-align:top">
      <div style="font-size:11px;color:${RC.muted}">心身のゆとり度</div>
      <div style="font-size:22px;font-weight:bold">${r.yutori}<span style="font-size:11px;color:${RC.muted}"> /100</span>　${pill_(['ゆとりあり','ややお疲れ気味','休息が必要'][yB[2]], yB[1])}</div>
      <div style="font-size:11px;color:${RC.muted};margin-top:6px">競技コンディション</div>
      <div style="font-size:22px;font-weight:bold">${r.condition}<span style="font-size:11px;color:${RC.muted}"> /100</span>　${pill_(cB[0], cB[1])}</div></td></tr></table>`;

  // 10カテゴリ
  h += h2_('10カテゴリのバランス（高いほど好調）');
  h += `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse">`;
  Object.keys(CATS).forEach(c=>{
    const v = r.s[c], b = band(v);
    h += `<tr><td style="padding:6px 6px 2px 0;font-size:12px;font-weight:bold;white-space:nowrap;width:34%">${MARK[c]} ${esc_(CATS[c].name)}</td>
      <td style="padding:6px 6px 2px;width:36%;vertical-align:middle">${bar_(v, b[1])}</td>
      <td style="padding:6px 6px 2px;font-size:12px;text-align:right;width:8%">${v}</td>
      <td style="padding:6px 0 2px 6px;width:22%">${pill_(b[0], b[1])}</td></tr>
      <tr><td colspan="4" style="padding:0 0 6px;font-size:11px;color:${RC.muted};line-height:1.6;border-bottom:1px solid ${RC.soft}">${esc_(COMMENT[c][b[2]])}</td></tr>`;
  });
  h += `</table>`;

  // 強み・伸びしろ
  h += h2_('あなたの強みと伸びしろ');
  h += `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse"><tr>
    <td style="width:50%;vertical-align:top;padding-right:10px"><div style="font-size:12px;font-weight:bold;color:${RC.good}">強み</div>
      ${r.strengths.map(([c,v])=>p(`${MARK[c]} ${esc_(CATS[c].name)}　${v}`)).join('')}${small(esc_(COMMENT[r.strengths[0][0]][0]))}</td>
    <td style="width:50%;vertical-align:top;padding-left:10px"><div style="font-size:12px;font-weight:bold;color:${RC.accent}">伸びしろ</div>
      ${r.growth.map(([c,v])=>p(`${MARK[c]} ${esc_(CATS[c].name)}　${v}`)).join('')}${small('ここを整えると、全体のコンディションが上がりやすいカテゴリです。')}</td></tr></table>`;

  // メッセージ
  h += h2_('あなたへのメッセージ');
  const msgs = r.chosen.length ? r.chosen : [{base:'自己肯定感', msg:COMMENT[r.strengths[0][0]][0]+'。それはあなたが積み重ねてきた力です。', ev:null}];
  msgs.forEach(m=>{
    h += `<div style="border-left:3px solid ${RC.accent};padding:2px 0 2px 12px;margin:8px 0">
      <div style="font-size:10px;font-weight:bold;color:${RC.accent}">${esc_(m.base)}</div>${p(esc_(m.msg))}
      ${m.ev===null ? small(`いちばん高かったのは「${esc_(CATS[r.strengths[0][0]].name)}」（${r.strengths[0][1]}点）でした。`) : small(esc_(evText_(m, a)))}</div>`;
  });

  if(!r.safety){
    h += h2_('あなたの「できる」の土台と、心のエネルギー源（推定）');
    h += `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse"><tr>
      <td style="width:46%;vertical-align:top;padding-right:10px"><div style="font-size:16px;font-weight:bold">${effWord}</div>
        ${small(r.efficacy>=70?'課題への見通しや立ち直った経験が、自信の根拠になっています。':r.efficacy>=50?'できることは確実に増えています。小さな「できた」を言葉にすると、さらに育ちます。':'自信は、過去にできたことを思い出すところから育て直せます。')}</td>
      <td style="vertical-align:top;padding-left:10px">
        <table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse">${Object.keys(r.sdt).map(k=>`<tr>
          <td style="font-size:11px;padding:3px 6px 3px 0;width:22%;${k===sdtMin?'font-weight:bold;color:'+RC.accent:''}">${k}</td>
          <td style="padding:3px 0">${bar_(r.sdt[k], k===sdtMin?'accent':'muted')}</td>
          <td style="font-size:11px;text-align:right;width:12%">${r.sdt[k]}</td></tr>`).join('')}</table>
        ${p(`いま満たしたいのは<b>「${sdtMin}」</b>。${{自律性:'自分で選べる部分を一つ増やしましょう。',有能感:'できたことを一つ記録しましょう。',関係性:'話せる人と過ごす時間を一つ増やしましょう。'}[sdtMin]}`)}</td></tr></table>`;
  }

  h += h2_('ワンポイント');
  opCats.forEach(c=>{ h += `<div style="margin:6px 0"><div style="font-size:12px;font-weight:bold">${MARK[c]} ${esc_(CATS[c].name)}（${r.s[c]}点）</div>${p(esc_(ONEPOINT[c][0]))}${small('自信の源：'+esc_(ONEPOINT[c][1]))}</div>`; });
  if(!r.safety) h += `<div style="background:${RC.soft};padding:10px 12px;margin-top:8px">${small(esc_(r.type)+'のあなたへ')}${p(esc_(T.msg))}</div>`;

  h += h2_('今日からできるセルフケア');
  h += careList.map((t,i)=>p(`${i+1}．${esc_(t)}`)).join('');

  if(!r.safety && r.s[4]<70){
    h += h2_('コーチ・監督との対話のヒント');
    h += small('いきなり不満を伝えるより、「確認」から入ると話しやすくなります。') + p('「今の練習で、自分が一番意識すべきことは何ですか？」');
  }
  if(!r.safety && (r.burnout!=='なし' || r.slump || r.motivTilt)){
    h += h2_('整えるとよいポイント');
    if(r.burnout!=='なし') h += p('疲れがたまり、手応えを感じにくくなっているようです。休養と練習量を見直すと、手応えが戻ってきます。');
    if(r.slump) h += p('試合での発揮度と、課題・意欲・心の状態をあわせて整えると、力を出しやすくなります。面談で一緒に整理することもできます。');
    if(r.motivTilt) h += p('「周りのため」に頑張れるのはあなたの良さです。「自分のため」の目標も一つ加えると、さらに力が湧いてきます。');
  }
  if(sa.S2==='yes'){
    h += h2_('指導者・チームとの関わりについて');
    h += p('怖い・つらいと感じる関わりがあると回答されました。あなたが我慢する必要はありません。チームや競技団体から独立した相談窓口に、匿名で相談することもできます。') + small('この回答はアスユメLabo.が管理し、チームや学校には共有されません。');
  }

  h += `<div style="margin-top:22px;padding-top:10px;border-top:1px solid ${RC.line}">
    ${small('このシートは医学的な診断ではありません。心や体の疲れが続くときは、医療機関や公認心理師などの専門家に相談すると回復の近道になります。')}
    ${small(TEL)}
    ${small('結果をもとに専門家と話したい方は、アスユメLabo.の面談をご利用ください。')}</div>`;
  h += `</div></body></html>`;
  return h;
}

/** メール本文（要点だけ。詳細は添付PDF） */
function buildMailHtml_(r, attrs){
  const yB = band(r.yutori), cB = band(r.condition);
  const first = r.chosen[0];
  const p = s => `<p style="margin:8px 0;font-size:14px;line-height:1.8;color:${RC.ink}">${s}</p>`;
  return `<div style="max-width:600px;margin:0 auto;font-family:${FONT};color:${RC.ink}">
    ${header_('cid:logo', null)}
    <div style="padding:20px 18px">
      ${p(esc_(attrs.name)+' さん')}
      ${p('KAKERU「わたしのいまここ」に回答いただき、ありがとうございました。フィードバックシートをお送りします。')}
      <table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;border:1px solid ${RC.line};margin:14px 0"><tr>
        <td style="padding:12px 14px;border-right:1px solid ${RC.line}"><div style="font-size:12px;color:${RC.muted}">あなたのタイプ</div><div style="font-size:20px;font-weight:bold">${esc_(r.type)}</div></td>
        <td style="padding:12px 14px;border-right:1px solid ${RC.line}"><div style="font-size:12px;color:${RC.muted}">心身のゆとり度</div><div style="font-size:20px;font-weight:bold">${r.yutori}</div>${pill_(['ゆとりあり','ややお疲れ気味','休息が必要'][yB[2]], yB[1])}</td>
        <td style="padding:12px 14px"><div style="font-size:12px;color:${RC.muted}">競技コンディション</div><div style="font-size:20px;font-weight:bold">${r.condition}</div>${pill_(cB[0], cB[1])}</td></tr></table>
      ${first ? `<div style="border-left:3px solid ${RC.accent};padding:2px 0 2px 12px">${p(esc_(first.msg))}</div>` : ''}
      ${p('10カテゴリの詳しい読み解き、ワンポイント、セルフケアは、添付のPDFをご覧ください。')}
      ${r.safety ? `<div style="border:1px solid ${RC.line};padding:10px 14px;margin:14px 0">${p('<b>よかったら、誰かに話してみませんか。</b>家族や友人、チームドクターなど、話しやすい人に声をかけてみてください。')}<p style="margin:4px 0;font-size:12px;color:${RC.muted}">${TEL}</p></div>` : ''}
      <p style="margin:18px 0 0;font-size:12px;line-height:1.7;color:${RC.muted}">受検ID：${esc_(r.id)}<br>このメールは送信専用ではありません。ご質問はこのメールにご返信ください。<br>このシートは医学的な診断ではありません。回答と結果はアスユメLabo.が管理し、チーム・学校・競技団体には共有しません。</p>
    </div></div>`;
}


/** アスユメLabo. ロゴ（PNG・背景透明）。画面ヘッダーとメール・PDFで使います。 */
const LOGO_B64 = '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';
