// Booking.gs - lets applicants pick who to talk to and a call time inside the page, using Google Calendar. No Calendly needed.
// Add as a new Script file named Booking (replaces the earlier Booking.gs).
var TZ = 'Asia/Manila';
var BOOK = {
  days: 7,          // how many days ahead applicants can book
  startHour: 9,     // first call starts at 9:00 AM
  endHour: 17,      // last call ends by 5:00 PM
  minutes: 15,      // length of each call
  leadHours: 2,     // earliest slot is this many hours from now
  daysOff: [7]      // 1 = Monday ... 7 = Sunday. [7] means no calls on Sunday
};
// cal: null = the calendar of the account that deploys this script (Manny).
// cal: '' = turned off. To turn on Annalyn: paste her Gmail below, after she shares her Google Calendar with Manny's account ("Make changes to events").
var PEOPLE = [
  {id: 'manny', name: 'Manny', cal: null},
  {id: 'annalyn', name: 'Annalyn', cal: 'anniemazing2@gmail.com'}
];

function calOf(p) { return p.cal ? CalendarApp.getCalendarById(p.cal) : CalendarApp.getDefaultCalendar(); }
function active() { return PEOPLE.filter(function (p) { return p.cal !== '' && calOf(p); }); }
function pick(who) { var a = active(); return (!who || who === 'any') ? a : a.filter(function (p) { return p.id === who; }); }
function getPeople() { return active().map(function (p) { return {id: p.id, name: p.name}; }); }

// Returns open times: [{day: '2026-10-03', times: [ISO strings]}]. A time is open if the chosen person (or anyone, for "any") is free.
function getSlots(who) {
  var now = new Date(), out = [];
  var busy = pick(who).map(function (p) {
    return calOf(p).getEvents(now, new Date(now.getTime() + (BOOK.days + 2) * 864e5))
      .filter(function (e) { return !e.isAllDayEvent(); })
      .map(function (e) { return [e.getStartTime().getTime(), e.getEndTime().getTime()]; });
  });
  if (!busy.length) return [];
  for (var i = 0; i <= BOOK.days; i++) {
    var d = new Date(now.getTime() + i * 864e5), key = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    if (BOOK.daysOff.indexOf(Number(Utilities.formatDate(d, TZ, 'u'))) > -1) continue;
    var times = [];
    for (var m = BOOK.startHour * 60; m + BOOK.minutes <= BOOK.endHour * 60; m += BOOK.minutes) {
      var hh = ('0' + Math.floor(m / 60)).slice(-2), mm = ('0' + (m % 60)).slice(-2);
      var st = Utilities.parseDate(key + ' ' + hh + ':' + mm, TZ, 'yyyy-MM-dd HH:mm').getTime(), en = st + BOOK.minutes * 6e4;
      if (st < now.getTime() + BOOK.leadHours * 36e5) continue;
      var free = busy.some(function (list) { return !list.some(function (b) { return st < b[1] && en > b[0]; }); });
      if (free) times.push(new Date(st).toISOString());
    }
    if (times.length) out.push({day: key, times: times});
  }
  return out;
}

// Books the chosen time on the chosen person's calendar (first free person for "any"), invites the applicant, marks the lead Scheduled.
function bookSlot(d) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var st = new Date(d.iso), en = new Date(st.getTime() + BOOK.minutes * 6e4);
    if (st.getTime() < Date.now()) return {ok: false, error: 'taken'};
    var people = pick(d.who), p = null, cal = null;
    for (var i = 0; i < people.length; i++) {
      var c = calOf(people[i]);
      if (!c.getEvents(st, en).filter(function (e) { return !e.isAllDayEvent(); }).length) { p = people[i]; cal = c; break; }
    }
    if (!p) return {ok: false, error: 'taken'};
    var desc = 'Home Partners PH recruitment call with ' + p.name + '\nName: ' + d.name + '\nMobile: ' + d.phone + '\nEmail: ' + d.email;
    try { cal.createEvent('Agent call with ' + p.name + ': ' + d.name, st, en, {description: desc, guests: d.email, sendInvites: true}); }
    catch (e) { cal.createEvent('Agent call with ' + p.name + ': ' + d.name, st, en, {description: desc}); }
    markScheduled(d.email, st, p.name);
    var when = Utilities.formatDate(st, TZ, 'EEE, MMM d, h:mm a');
    try {
      MailApp.sendEmail(NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(), 'Call booked with ' + p.name + ': ' + d.name + ' (' + when + ')',
        'Call booked with ' + p.name + '\nName: ' + d.name + '\nTime: ' + when + ' (Philippine time)\nMobile: ' + d.phone + '\nEmail: ' + d.email);
    } catch (e) {}
    return {ok: true, when: when, with: p.name};
  } finally { lock.releaseLock(); }
}

function markScheduled(email, st, name) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0], n = sh.getLastRow();
  if (n < 2) return;
  var emails = sh.getRange(2, 4, n - 1, 1).getValues();
  for (var i = emails.length - 1; i >= 0; i--) {
    if (String(emails[i][0]).toLowerCase() === String(email).toLowerCase()) {
      sh.getRange(i + 2, 11).setValue('Scheduled');
      sh.getRange(i + 2, 14).setValue(Utilities.formatDate(st, TZ, 'yyyy-MM-dd HH:mm') + ' with ' + name);
      return;
    }
  }
}

// Run once from the editor to allow Calendar access and check the people and slots
function testBooking() {
  Logger.log(JSON.stringify(getPeople()));
  Logger.log(JSON.stringify(getSlots('any').slice(0, 2)));
}
