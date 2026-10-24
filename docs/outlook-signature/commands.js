/*
 * JHC Outlook signature add-in: event handler.
 *
 * Runs when Outlook opens a new message, reply, reply-all or forward
 * (the OnNewMessageCompose launch event). It reads the signed-in user's
 * email, looks it up in signatures.json, and inserts that person's
 * signature. If the email is not in the file, or anything fails, it
 * inserts nothing and lets Outlook carry on.
 *
 * Rules for this file (classic Outlook on Windows runs it in a
 * JavaScript-only runtime): one file, no import/require, no window or
 * document, absolute fetch URLs only, and event.completed() on every path.
 * Full doc: README.md in the repo root.
 */

var JHC_SIG = (function () {
  var SIGNATURES_URL = "https://justinhall02.github.io/outlook-signature/signatures.json";
  var FETCH_TIMEOUT_MS = 10000;

  var HOME_URL = "https://justinhallconsulting.com/";
  var ALARM_URL = "https://www.alarm.com/US/JustinHallConsulting?home=1";
  var LOGO_URL = "https://justinhallconsulting.com/wp-content/uploads/2026/10/brandmark-design-av.png";
  var ALARM_LOGO_URL = "https://justinhallconsulting.com/wp-content/uploads/2025/09/Alarm_logo.png";
  var MAIN_PHONE = "404-905-2213";
  var SERVICE_EMAIL = "Service@JustinHallConsulting.com";

  var FONT = "font-family:Arial,Helvetica,sans-serif;";
  var LINK = "color:rgb(10,102,194); text-decoration:none;";

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function safeUrl(u) {
    return /^https:\/\//i.test(String(u || "")) ? esc(u) : "";
  }

  // Same layout as the signatures built 2026-10-22 (signatures-rollout),
  // with the Outlook-on-the-web clutter and HubSpot tracking links removed.
  function render(p) {
    var phone = MAIN_PHONE + (p.ext ? " ext. " + esc(p.ext) : "");
    var left = "";
    left += '<div style="line-height:1.2; margin:0px; ' + FONT + ' font-size:18px; color:rgb(15,23,42)"><b>' + esc(p.name) + "</b></div>";
    if (p.title) {
      left += '<div style="line-height:1.4; margin:2px 0px 10px; ' + FONT + ' font-size:13px; color:rgb(51,65,85)">' + esc(p.title) + "</div>";
    } else {
      left += '<div style="line-height:1.4; margin:0px 0px 10px; font-size:1px">&nbsp;</div>';
    }
    left += '<div style="line-height:1.6; margin:0px; ' + FONT + ' font-size:13px"><span style="color:rgb(31,41,55)"><b>Office:</b>&nbsp;</span><span style="color:rgb(10,102,194)">' + phone + "</span></div>";
    left += '<div style="line-height:1.6; margin:0px; ' + FONT + ' font-size:13px"><span style="color:rgb(31,41,55)"><br><b>Service Requests:</b>&nbsp;</span><a href="mailto:' + SERVICE_EMAIL + '" style="' + LINK + '">' + SERVICE_EMAIL + "</a></div>";
    var booking = safeUrl(p.bookingUrl);
    var review = safeUrl(p.reviewUrl);
    if (booking || review) {
      left += '<div style="line-height:1.6; margin:2px 0px 0px; ' + FONT + ' font-size:13px">';
      if (booking) left += '<a href="' + booking + '" style="' + LINK + '"><b>Book time with me</b></a>';
      if (booking && review) left += "<br>";
      if (review) left += '<a href="' + review + '" style="' + LINK + '"><b>Share your feedback on Google</b></a>';
      left += "</div>";
    }

    var right = "";
    right += '<div style="margin:0px"><a href="' + HOME_URL + '" style="text-decoration:none"><img alt="Justin Hall Consulting" width="230" src="' + LOGO_URL + '" style="width:230px; max-width:230px; margin:0px; display:block; border:0"></a></div>';
    right += '<div style="margin:0px; ' + FONT + ' font-size:12px; color:rgb(51,65,85)">&nbsp; &nbsp; &nbsp; &nbsp; We Construct Business Technology</div>';
    right += '<div style="margin:0px; ' + FONT + ' font-size:11px; color:rgb(71,85,105)">Cabling &middot; Security &middot; Telecom &amp; Mobile &middot; AV</div>';
    right += '<div style="margin:0px; ' + FONT + ' font-size:11px; line-height:28px">&nbsp;</div>';
    right += '<div style="margin:0px"><a href="' + ALARM_URL + '" style="text-decoration:none"><img alt="Authorized Alarm.com Partner" width="146" height="63" src="' + ALARM_LOGO_URL + '" style="width:146px; height:63px; margin:0px; display:block; border:0"></a></div>';
    right += '<div style="line-height:1.4; margin:0px; ' + FONT + ' font-size:12px; color:rgb(31,41,55)"><b>W:</b>&nbsp;<a href="' + HOME_URL + '" style="' + LINK + '">justinhallconsulting.com</a></div>';

    return (
      '<table cellspacing="0" cellpadding="0" style="background-color:rgb(255,255,255); width:613px; color:rgb(31,41,55); border-collapse:collapse; border-spacing:0px"><tbody><tr>' +
      '<td style="border-right:1px solid rgb(209,213,219); padding:12px 16px 12px 0px; vertical-align:top; width:340px">' + left + "</td>" +
      '<td style="padding:12px 0px 12px 16px; vertical-align:top; width:260px">' + right + "</td>" +
      "</tr></tbody></table>"
    );
  }

  function findPerson(data, email) {
    var want = String(email || "").trim().toLowerCase();
    if (!want || !data || !data.people) return null;
    for (var i = 0; i < data.people.length; i++) {
      var p = data.people[i];
      var all = [p.email].concat(p.aliases || []);
      for (var j = 0; j < all.length; j++) {
        if (String(all[j] || "").trim().toLowerCase() === want) return p;
      }
    }
    return null;
  }

  function loadSignatures() {
    var timeout = new Promise(function (_, reject) {
      setTimeout(function () { reject(new Error("timeout")); }, FETCH_TIMEOUT_MS);
    });
    var req = fetch(SIGNATURES_URL, { cache: "no-store" }).then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.json();
    });
    return Promise.race([req, timeout]);
  }

  function supports(set) {
    try {
      return Office.context.requirements.isSetSupported("Mailbox", set);
    } catch (e) {
      return false;
    }
  }

  function setSignature(item, html, done) {
    item.body.setSignatureAsync(html, { coercionType: Office.CoercionType.Html }, function () { done(); });
  }

  function onNewMessageComposeHandler(event) {
    var finished = false;
    function done() {
      if (finished) return;
      finished = true;
      try { event.completed(); } catch (e) { /* nothing else to do */ }
    }
    try {
      var item = Office.context.mailbox.item;
      var email = Office.context.mailbox.userProfile.emailAddress;
      loadSignatures()
        .then(function (data) {
          var person = findPerson(data, email);
          if (!person) return done();
          var html = render(person);
          // Turn off Outlook's own signature on this message so there are not two.
          if (item.isClientSignatureEnabledAsync && item.disableClientSignatureAsync && supports("1.10")) {
            item.isClientSignatureEnabledAsync(function (r) {
              if (r.status === Office.AsyncResultStatus.Succeeded && r.value === true) {
                item.disableClientSignatureAsync(function () { setSignature(item, html, done); });
              } else {
                setSignature(item, html, done);
              }
            });
          } else {
            setSignature(item, html, done);
          }
        })
        .catch(function () { done(); });
    } catch (e) {
      done();
    }
  }

  return { onNewMessageComposeHandler: onNewMessageComposeHandler, render: render, findPerson: findPerson };
})();

function onNewMessageComposeHandler(event) {
  return JHC_SIG.onNewMessageComposeHandler(event);
}

if (typeof Office !== "undefined") {
  if (Office.onReady) Office.onReady();
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  }
}
