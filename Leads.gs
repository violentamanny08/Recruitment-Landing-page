// Leads.gs - saves leads, emails the applicant, sends a 30-day short-line conversational follow-up sequence, tracks source and status.
// Replace the whole Leads.gs file with this. Needs Email.gs (esc, sendLeadEmail, EMAIL_PHOTO, EMAIL_LOGO) and SHEET_ID / NOTIFY_EMAIL from Code.gs.
var BOOKING_URL = '';  // paste your call-booking link here (Google Calendar appointment page, Calendly, etc.). Blank = Call button only.
var STATUS_LIST = ['New', 'Called', 'Scheduled', 'Onboarded', 'Not a fit', 'Stop emails'];
var HEADERS = ['Time','Name','Mobile','Email','City','Score','Answers','Source','Language','Tier','Status','FollowUpStage','LastFollowUp','CallTime'];

// Day (counted from the day they applied) when each follow-up goes out. 8 emails across 30 days.
var FOLLOW_DAYS = [1, 3, 6, 10, 14, 19, 24, 30];
var KINDS = ['hot', 'f1', 'f2', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8'];

function tierOf(s) { return s >= 75 ? 'Strong fit' : s >= 50 ? 'Could be a fit' : 'Not yet'; }

// Called by the page when someone finishes the quiz
function submitLead(d) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sh.getRange('C:C').setNumberFormat('@');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  var score = Number(d.score) || 0, tier = tierOf(score);
  sh.appendRow([new Date(), d.name, d.phone, d.email, d.city, score, (d.answers || []).join(' | '),
                d.source || 'direct', d.language || 'en', tier, 'New', 0, '']);
  sh.getRange('K2:K1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS_LIST, true).setAllowInvalid(false).build());
  sendLeadEmail(d);
  try { sendApplicantEmail(d, tier, 0); } catch (e) {}
  return 'ok';
}

var COPY = {
  // Day 0: confirmation
  hot: {
    en: {s: 'We got your application, {n}', h: 'Hi {n}, it is Manny and Annalyn!',
         p: ['Thank you for applying to join our team at Home Partners PH. We got your answers.',
             'We will read your application ourselves.',
             'Expect a call or a message from us soon, so we can get to know you and answer your questions.',
             'You can also reply to this email, or use the button below.',
             'We are excited to meet you!',
             '- Manny and Annalyn']},
    tl: {s: 'Natanggap na namin ang application mo, {n}', h: 'Hi {n}, Manny at Annalyn here!',
         p: ['Salamat sa pag-apply sa team namin sa Home Partners PH. Natanggap na namin ang mga sagot mo.',
             'Kami mismo ang magbabasa ng application mo.',
             'Tatawagan o ime-message ka namin soon para makilala ka namin at masagot ang mga tanong mo.',
             'Pwede ka ring mag-reply sa email na ito, o gamitin ang button sa ibaba.',
             'Excited kaming makilala ka!',
             '- Manny at Annalyn']}},

  // Day 1: quick hello
  f1: {
    en: {s: 'Quick hello, {n}', h: 'Hi {n}, how are you?',
         p: ['It\'s Manny and Annalyn.',
             'You sent in your application yesterday.',
             'So we just wanted to say hi again.',
             'Maybe you\'re wondering, \"What happens next?\"',
             'Good news. It\'s really simple.',
             'Just a short call. About 15 minutes.',
             'It\'s not a scary interview.',
             'It\'s just a conversation.',
             'We\'ll share how we work.',
             'You tell us why you wanted to apply.',
             'Got questions before we talk?',
             'Even the big ones, like \"How much can I earn?\"',
             'Or \"Can I really do this?\"',
             'We\'ll answer honestly.',
             'Just reply to this email.',
             'Or tap the button below.',
             'Looking forward to hearing from you, {n}.',
             '- Manny and Annalyn']},
    tl: {s: 'Kumusta, {n}? Usap tayo?', h: 'Hi {n}, kumusta ka na?',
         p: ['Si Manny at Annalyn ito.',
             'Kahapon mo lang sinagutan ang application mo.',
             'Kaya gusto lang naming mag-hi ulit.',
             'Baka iniisip mo, \"Ano na kaya ang susunod?\"',
             'Simple lang, promise.',
             'Isang maikling tawag lang. Mga 15 minutes.',
             'Hindi ito nakakakabang interview.',
             'Kwentuhan lang talaga.',
             'Ikukwento namin kung paano kami magtrabaho.',
             'Ikaw naman, ikwento mo kung bakit ka nag-apply.',
             'May tanong ka ba bago tayo mag-usap?',
             'Kahit \"Magkano ang pwede kong kitain?\"',
             'O \"Kaya ko ba talaga ito?\"',
             'Sasagutin namin nang tapat.',
             'Mag-reply ka lang dito.',
             'O pindutin ang button sa ibaba.',
             'Hihintayin ka namin, {n}.',
             '- Manny at Annalyn']}},

  // Day 3: what the work and training look like
  f2: {
    en: {s: 'What will you actually do, {n}?', h: 'Hi {n}, here is what to expect',
         p: ['It\'s Manny and Annalyn again.',
             'One question we hear a lot:',
             '\"What will I actually do if I join?\"',
             'So let us tell you straight.',
             'First, we teach you the basics.',
             'How to list a property.',
             'How to price it right.',
             'How to negotiate, and how to close.',
             'It\'s not a one-day crash course.',
             'It\'s step by step, at your pace.',
             'Second, you\'re never alone.',
             'You\'ll have real people to ask when you\'re lost.',
             'Third, we\'ll show you how to market listings.',
             'And follow up buyers on Facebook and TikTok.',
             'You don\'t have to be good on day one.',
             'You just have to be willing to learn.',
             'Got a question already?',
             'Just reply and ask.',
             'We\'d love to hear from you, {n}.',
             '- Manny and Annalyn']},
    tl: {s: 'Ano ba talaga ang gagawin mo, {n}?', h: 'Hi {n}, ito ang aasahan mo kapag sumali ka',
         p: ['Manny at Annalyn ulit.',
             'May tanong na madalas naming marinig.',
             '\"Ano ba talaga ang gagawin ko kapag sumali ako?\"',
             'Sasabihin na namin nang diretso.',
             'Una, tuturuan ka namin ng basics.',
             'Paano mag-list ng property.',
             'Paano mag-presyo nang tama.',
             'Paano makipag-negotiate, at paano mag-close.',
             'Hindi ito pang-isang araw lang.',
             'Step by step, sa bilis na kaya mo.',
             'Pangalawa, hindi ka mag-isa.',
             'May makakausap ka kapag nalilito ka.',
             'Pangatlo, tuturuan ka naming mag-market ng listings.',
             'At mag-follow up ng buyers sa Facebook at TikTok.',
             'Hindi mo kailangang magaling agad.',
             'Ang kailangan lang, handa kang matuto.',
             'May tanong ka na ba ngayon?',
             'Itanong mo lang. Mag-reply ka dito.',
             'Excited kaming marinig mula sa iyo, {n}.',
             '- Manny at Annalyn']}},

  // Day 6: pain point - the paycheck is not enough
  f3: {
    en: {s: 'Is your paycheck enough, {n}?', h: 'Hi {n}, can we be real for a minute?',
         p: ['Manny and Annalyn again.',
             'We have a question.',
             'We hope it\'s okay to ask.',
             'Is what you earn right now enough?',
             'We\'re not judging you.',
             'It\'s just that many people apply for this reason.',
             'The salary is gone before the month ends.',
             'You\'re tired from work.',
             'But there\'s nothing left for savings.',
             'For family.',
             'Or for your dreams.',
             'We get it.',
             'And we want to be honest with you.',
             'Real estate is not magic.',
             'You won\'t get rich overnight.',
             'But your effort directly grows what you earn.',
             'Want to know how it works?',
             'Reply with \"Tell me more\".',
             'We\'ll explain it clearly. No pressure.',
             'Waiting to hear from you, {n}.',
             '- Manny and Annalyn']},
    tl: {s: 'Sapat ba ang sweldo mo, {n}?', h: 'Hi {n}, pwede bang maging totoo muna tayo saglit?',
         p: ['Manny at Annalyn ito ulit.',
             'May gusto lang kaming itanong.',
             'Sana okay lang sa iyo.',
             'Sapat ba ang kinikita mo ngayon?',
             'Hindi kami nanghuhusga.',
             'Kasi madalas, ito ang dahilan ng mga nag-a-apply sa amin.',
             'Ubos na ang sweldo bago matapos ang buwan.',
             'Pagod ka na sa trabaho.',
             'Pero walang natitira para sa ipon.',
             'Para sa pamilya.',
             'O para sa mga pangarap mo.',
             'Naiintindihan namin iyon.',
             'At gusto naming maging tapat sa iyo.',
             'Hindi magic ang real estate.',
             'Hindi ka yayaman sa isang gabi.',
             'Pero ang sipag at diskarte mo mismo ang magpapalaki ng kita mo.',
             'Gusto mong malaman kung paano?',
             'Mag-reply ka lang ng \"Sabihin mo pa\".',
             'Ipapaliwanag namin nang malinaw. Walang pilitan.',
             'Hinihintay namin ang sagot mo, {n}.',
             '- Manny at Annalyn']}},

  // Day 10: pain point - no experience, fear of not being good enough
  f4: {
    en: {s: 'No experience? That\'s okay, {n}', h: 'Hi {n}, about that \"what if I am not good enough\" feeling',
         p: ['It\'s Manny and Annalyn.',
             'Has a little voice in your head ever said this?',
             '\"Sounds great, but I have no experience.\"',
             'If yes, you\'re not alone.',
             'Many people start knowing very little.',
             'No clients yet.',
             'Sometimes not much confidence either.',
             'And that\'s okay.',
             'Experience isn\'t the most important thing at the start.',
             'Being willing to learn is.',
             'Being willing to talk to people.',
             'And being willing to follow up.',
             'We\'ll handle the guidance.',
             'From your first listing to your first closing.',
             'You won\'t be guessing.',
             'You won\'t be alone.',
             'What holds most people back is \"What if I can\'t?\"',
             'You won\'t know until you try.',
             'And we\'ll be right here with you.',
             'Want to ask or share something? Just reply.',
             'We\'re rooting for you, {n}.',
             '- Manny and Annalyn']},
    tl: {s: 'Walang experience? Okay lang iyan, {n}', h: 'Hi {n}, tungkol sa takot na \"baka hindi ako sapat\"',
         p: ['Si Manny at Annalyn ito.',
             'May bumulong na ba sa isip mo na ganito?',
             '\"Maganda ang idea, pero wala akong experience.\"',
             'Kung oo, hindi ka nag-iisa.',
             'Marami ang nagsisimula nang wala pang alam.',
             'Wala pang client.',
             'Minsan wala pang kumpiyansa.',
             'At okay lang iyon.',
             'Hindi experience ang pinakamahalaga sa simula.',
             'Ang mahalaga, handa kang matuto.',
             'Handa kang makipag-usap sa tao.',
             'At handa kang mag-follow up.',
             'Kami na ang bahala sa gabay.',
             'Mula sa unang listing hanggang sa unang closing.',
             'Hindi ka manghuhula.',
             'Hindi ka mag-isa.',
             'Ang pumipigil sa marami ay ang \"Paano kung hindi ko kaya?\"',
             'Hindi mo malalaman hangga\'t hindi mo sinusubukan.',
             'At nandito kami para samahan ka.',
             'May gusto ka bang itanong? Mag-reply ka lang.',
             'Cheer kami sa iyo, {n}.',
             '- Manny at Annalyn']}},

  // Day 14: pain point - where will my clients come from
  f5: {
    en: {s: 'Where do the clients come from, {n}?', h: 'Hi {n}, this is everyone\'s biggest question',
         p: ['Manny and Annalyn again.',
             'Know what we hear most?',
             '\"Where will my clients come from?\"',
             'It\'s a fair question.',
             'Many new agents give up.',
             'Not because they\'re weak.',
             'But because they try to figure it out alone.',
             'They post randomly.',
             'They wait for someone to message.',
             'Then they slowly lose energy.',
             'We don\'t want that for you.',
             'So we\'ll teach you a simple system.',
             'How to market listings on Facebook and TikTok.',
             'How to answer inquiries fast.',
             'How to follow up without being pushy.',
             'And when you\'re stuck, you can message us.',
             'Want to see what a normal day looks like?',
             'Reply \"Show me\".',
             'All you have to do is listen, {n}.',
             '- Manny and Annalyn']},
    tl: {s: 'Saan nanggagaling ang mga client, {n}?', h: 'Hi {n}, ito ang pinakamalaking tanong ng lahat',
         p: ['Manny at Annalyn ulit.',
             'Alam mo ba ang pinakamadalas naming marinig?',
             '\"Saan naman manggagaling ang mga client ko?\"',
             'Tama lang na itanong mo iyan.',
             'Maraming bagong agent ang sumusuko.',
             'Hindi dahil mahina sila.',
             'Kundi dahil mag-isa nilang inaalam ang lahat.',
             'Nagpo-post nang random.',
             'Naghihintay na may mag-message.',
             'Tapos unti-unting nawawalan ng gana.',
             'Ayaw naming mangyari iyon sa iyo.',
             'Kaya may ituturo kaming simpleng sistema.',
             'Paano mag-market sa Facebook at TikTok.',
             'Paano sumagot agad sa mga nagtatanong.',
             'Paano mag-follow up nang hindi mapilit.',
             'At kapag na-stuck ka, may dalawang taong pwede mong i-message.',
             'Kami iyon.',
             'Gusto mong makita ang normal na araw ng agent sa team namin?',
             'Mag-reply ka ng \"Ipakita mo\".',
             'Pakikinig lang ang kailangan mong gawin, {n}.',
             '- Manny at Annalyn']}},

  // Day 19: pain point - is this legit / commission fear
  f6: {
    en: {s: 'Is this legit, {n}? Fair question', h: 'Hi {n}, it\'s okay to have doubts',
         p: ['It\'s Manny and Annalyn.',
             'Maybe you\'re thinking this. Maybe not.',
             '\"Is this really legit?\"',
             'If yes, that\'s okay.',
             'Many people have been fooled by scams.',
             'So we understand if you\'re careful.',
             'So here\'s the straight answer.',
             'This is real estate.',
             'You earn commission when a deal closes.',
             'There\'s no guaranteed salary.',
             'And nobody should promise you one.',
             'What can we promise?',
             'Honest guidance.',
             'Real training.',
             'And people who answer your questions.',
             'Have doubts? Ask us.',
             'Even about earnings and costs. We\'ll answer straight.',
             'You can also get to know us first.',
             'The call is only 15 minutes.',
             'Our agent portal is linked below.',
             'There\'s nothing wrong with being careful, {n}.',
             'We actually prefer it.',
             '- Manny and Annalyn']},
    tl: {s: 'Legit ba ito, {n}? Tama lang na magtanong', h: 'Hi {n}, okay lang kung may duda ka',
         p: ['Manny at Annalyn ito.',
             'Baka iniisip mo ito, baka hindi.',
             '\"Legit ba talaga ito?\"',
             'Kung oo, okay lang.',
             'Maraming Pilipino ang naloko na ng scam.',
             'Kaya naiintindihan namin kung maingat ka.',
             'Kaya sasabihin namin nang diretso.',
             'Real estate ito.',
             'Kumikita ka ng commission kapag nag-close ang deal.',
             'Walang garantisadong sweldo.',
             'At walang dapat mangako sa iyo noon.',
             'Ang kaya naming ipangako?',
             'Tapat na gabay.',
             'Totoong training.',
             'At mga taong sasagot sa tanong mo.',
             'May duda ka ba? Itanong mo.',
             'Kahit tungkol sa kita at gastos, sasagutin namin nang diretso.',
             'Pwede mo rin kaming kilalanin muna.',
             '15 minutes lang ang tawag.',
             'Nasa ibaba ang link ng agent portal namin.',
             'Walang masama sa pag-iingat, {n}.',
             'Mas gusto pa nga namin iyon.',
             '- Manny at Annalyn']}},

  // Day 24: pain point - no time, no control of the schedule
  f7: {
    en: {s: 'More time for what matters, {n}', h: 'Hi {n}, try picturing this',
         p: ['Manny and Annalyn again.',
             'Picture your usual week.',
             'Wake up early.',
             'Commute. Work.',
             'Go home tired.',
             'You have a boss.',
             'A schedule you didn\'t choose.',
             'And hardly any time left for family.',
             'Sound familiar?',
             'Many of our applicants feel the same.',
             'It\'s not only about money.',
             'They want control over their time too.',
             'Real estate can give you that.',
             'But let\'s be honest.',
             'You have to build your skills and network first.',
             'It\'s not instant.',
             'It\'s real work.',
             'But the work is yours.',
             'You hold your hustle and your time.',
             'Want to give it a try?',
             'Reply \"Yes\".',
             'We\'ll find a time that fits your schedule, {n}.',
             '- Manny and Annalyn']},
    tl: {s: 'Mas maraming oras para sa mahalaga, {n}', h: 'Hi {n}, isipin mo muna ito',
         p: ['Si Manny at Annalyn ito ulit.',
             'Subukan mong isipin ang karaniwang linggo mo.',
             'Gising nang maaga.',
             'Biyahe. Trabaho.',
             'Uwi nang pagod.',
             'May boss ka.',
             'May schedule na hindi ikaw ang pumili.',
             'At pag-uwi mo, halos wala nang oras para sa pamilya.',
             'Pamilyar ba?',
             'Maraming applicant namin ang ganito.',
             'Hindi lang pera ang gusto nila.',
             'Gusto rin nila ng kontrol sa oras nila.',
             'Pwedeng ibigay iyon ng real estate.',
             'Pero maging tapat tayo.',
             'Kailangan mo munang buuin ang skills at network mo.',
             'Hindi ito instant.',
             'Totoong trabaho ito.',
             'Pero ang maganda, sa iyo ang trabahong iyon.',
             'Ikaw ang may hawak ng sipag at oras mo.',
             'Gusto mong subukan?',
             'Mag-reply ka ng \"Oo\".',
             'Maghahanap kami ng oras na bagay sa schedule mo, {n}.',
             '- Manny at Annalyn']}},

  // Day 30: last email, two clear choices
  f8: {
    en: {s: 'Our last note, {n}', h: 'Hi {n}, this is our last email',
         p: ['It\'s Manny and Annalyn.',
             'It\'s been about a month since you applied.',
             'We\'ve written a few times.',
             'We don\'t want to crowd your inbox.',
             'So this is the last one.',
             'We\'ll make your answer easy.',
             'Still interested?',
             'Just reply \"Yes\".',
             'We\'ll set up a call at a time that works for you.',
             'Not for you right now?',
             'Reply \"Pass\", or just don\'t reply.',
             'We won\'t send any more emails.',
             'No anger. No hurt feelings.',
             'Everyone\'s timing is different.',
             'Maybe you\'re busy.',
             'Maybe you\'re sorting things out.',
             'Maybe it\'s not the right season.',
             'If you change your mind someday, our door is open.',
             'Thank you for your time, {n}.',
             'Take care always.',
             '- Manny and Annalyn']},
    tl: {s: 'Huling mensahe na namin, {n}', h: 'Hi {n}, ito na ang huli naming email',
         p: ['Si Manny at Annalyn ito.',
             'Mga isang buwan na mula nung nag-apply ka.',
             'Ilang beses na kaming sumulat.',
             'Ayaw naming maging pabigat sa inbox mo.',
             'Kaya ito na ang huli.',
             'Gagawin naming madali ang sagot mo.',
             'Interesado ka pa rin?',
             'Mag-reply ka ng \"Oo\".',
             'Ia-arrange namin ang tawag sa oras na bagay sa iyo.',
             'Hindi na ba ito para sa iyo ngayon?',
             'Mag-reply ka ng \"Pass\", o huwag na lang sumagot.',
             'Hindi na kami magpapadala ng email.',
             'Walang galit. Walang tampo.',
             'Iba-iba ang timing ng bawat tao.',
             'Baka busy ka.',
             'Baka may inaayos ka.',
             'Baka hindi pa ito ang tamang panahon.',
             'Kung magbago ang isip mo balang araw, bukas pa rin ang pinto namin.',
             'Salamat sa oras mo, {n}.',
             'Ingat ka palagi.',
             '- Manny at Annalyn']}},

  // Applicants with tier "Not yet"
  no: {
    en: {s: 'Thank you for applying, {n}', h: 'Hi {n}, thank you for your time',
         p: ['We got your application, and thank you for your honest answers.',
             'Right now it looks like the timing is not quite right for this role, and that is okay.',
             'If you ever have more time or a steadier income runway, you are welcome to apply again.',
             'We hope you have a great day!',
             '- Manny and Annalyn']},
    tl: {s: 'Salamat sa pag-apply, {n}', h: 'Hi {n}, salamat sa oras mo',
         p: ['Natanggap namin ang application mo, at salamat sa tapat mong mga sagot.',
             'Sa ngayon, parang hindi pa ito ang tamang panahon para sa role na ito, at okay lang iyon.',
             'Kung magkaroon ka ng mas maraming oras o mas matatag na pagkukunan ng kita, malugod kang mag-apply ulit.',
             'Sana maging maganda ang araw mo!',
             '- Manny at Annalyn']}}
};

// stage 0 = confirmation, 1-8 = follow-ups. Applicants with tier "Not yet" only get the polite note.
function sendApplicantEmail(d, tier, stage) {
  var lang = d.language === 'tl' ? 'tl' : 'en';
  var kind = tier === 'Not yet' ? 'no' : KINDS[stage];
  var first = String(d.name || '').trim().split(' ')[0] || (lang === 'tl' ? 'kaibigan' : 'there');
  var m = COPY[kind][lang];
  var fill = function (t) { return t.replace(/\{n\}/g, first); };
  var heading = fill(m.h), subject = fill(m.s);
  var cta = '', stopNote = '';
  if (kind !== 'no') {
    var url = BOOKING_URL || 'tel:+639953581519';
    var label = BOOKING_URL ? (lang === 'tl' ? 'Mag-book ng tawag' : 'Book a call') : (lang === 'tl' ? 'Tawagan ang 0995 358 1519' : 'Call 0995 358 1519');
    cta = '<a href="' + url + '" style="display:inline-block;background:#ff6b00;color:#031a3f;font:bold 15px Arial,sans-serif;text-decoration:none;padding:13px 26px;border-radius:24px">' + label + '</a>';
    cta += ' <a href="https://innersparcagentportal.vercel.app" style="display:inline-block;background:#062b67;color:#ffffff;font:bold 15px Arial,sans-serif;text-decoration:none;padding:13px 26px;border-radius:24px;margin-top:8px">' +
           (lang === 'tl' ? 'Bisitahin ang agent portal' : 'Visit our agent portal') + '</a>';
  }
  if (kind.charAt(0) === 'f') {
    stopNote = '<p style="font:12px/1.5 Arial,sans-serif;color:#8a97a8;margin:18px 0 0">' +
               (lang === 'tl' ? 'Hindi ka na interesado? Mag-reply lang ng STOP at ihihinto namin ang mga email.' : 'Not interested anymore? Just reply STOP and we will stop the emails.') + '</p>';
  }
  var paras = m.p.map(function (t) { return '<p style="font:16px/1.6 Arial,sans-serif;color:#2b3f57;margin:0 0 10px">' + esc(t) + '</p>'; }).join('');
  var html =
    '<div style="background:#eef3fa;padding:24px 10px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:16px;overflow:hidden">' +
    '<tr><td style="background:#031a3f;padding:16px 24px"><img src="cid:logo" width="44" height="44" alt="" style="vertical-align:middle;border-radius:22px"> ' +
    '<span style="font:bold 18px Arial,sans-serif;color:#ffffff;vertical-align:middle;margin-left:10px">Home Partners PH</span></td></tr>' +
    '<tr><td><img src="cid:photo" width="600" alt="Manny and Annalyn" style="display:block;width:100%;max-width:600px;height:auto"></td></tr>' +
    '<tr><td style="background:#ff6b00;height:5px;font-size:0;line-height:0">&nbsp;</td></tr>' +
    '<tr><td style="padding:28px 24px 8px"><div style="font:bold 24px Arial,sans-serif;color:#062b67;margin:0 0 14px">' + esc(heading) + '</div>' + paras + '</td></tr>' +
    '<tr><td style="padding:6px 24px 26px">' + cta + stopNote + '</td></tr>' +
    '<tr><td style="background:#031a3f;padding:14px 24px;font:12px Arial,sans-serif;color:#b9c9de;text-align:center">Home Partners PH &middot; Real Life Partners &middot; Real Service &middot; Hassle-Free</td></tr>' +
    '</table></div>';
  MailApp.sendEmail({
    to: d.email, subject: subject, body: heading + '\n\n' + m.p.map(fill).join('\n\n'), htmlBody: html,
    name: 'Home Partners PH', replyTo: NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(),
    inlineImages: {
      photo: Utilities.newBlob(Utilities.base64Decode(EMAIL_PHOTO), 'image/jpeg', 'photo.jpg'),
      logo: Utilities.newBlob(Utilities.base64Decode(EMAIL_LOGO), 'image/jpeg', 'logo.jpg')
    }
  });
}

// Runs every day (see setupTriggers). Sends 8 follow-ups over 30 days (see FOLLOW_DAYS).
// Only leads with Status "New" and tier Strong fit or Could be a fit. Change Status (Scheduled, Called, Stop emails...) to stop the emails.
function followUps() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0], n = sh.getLastRow();
  if (n < 2) return;
  var rows = sh.getRange(2, 1, n - 1, 13).getValues(), now = new Date();
  rows.forEach(function (r, k) {
    var tier = r[9], status = r[10], stage = Number(r[11]) || 0;
    if (!r[3] || !tier || tier === 'Not yet' || status !== 'New' || stage >= FOLLOW_DAYS.length) return;
    if ((now - new Date(r[0])) / 86400000 < FOLLOW_DAYS[stage]) return;
    try {
      sendApplicantEmail({name: r[1], email: r[3], language: r[8]}, tier, stage + 1);
      sh.getRange(k + 2, 12, 1, 2).setValues([[stage + 1, now]]);
    } catch (e) {}
  });
}

// Run once from the editor: sets the daily 9am follow-up trigger, headers, and the Status dropdown.
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'followUps') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('followUps').timeBased().everyDays(1).atHour(9).create();
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.getRange('K2:K1000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUS_LIST, true).setAllowInvalid(false).build());
}

// Preview all applicant emails in English (sends 11 emails to NOTIFY_EMAIL, or the account owner)
function testApplicantEmails() {
  var me = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  for (var s = 0; s <= 8; s++) {
    sendApplicantEmail({name: 'Test Applicant', email: me, language: 'en'}, 'Strong fit', s);
  }
  sendApplicantEmail({name: 'Test Applicant', email: me, language: 'en'}, 'Not yet', 0);
  sendApplicantEmail({name: 'Test Applicant', email: me, language: 'tl'}, 'Strong fit', 3);
}

// Preview only the Tagalog follow-ups (sends 8 emails)
function testTagalogEmails() {
  var me = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  for (var s = 1; s <= 8; s++) {
    sendApplicantEmail({name: 'Test Applicant', email: me, language: 'tl'}, 'Strong fit', s);
  }
}
