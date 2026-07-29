/**
 * U.S. Stoneworks — booking logger
 * ---------------------------------
 * Receives booking submissions from the website's "Confirm Booking" step
 * and appends each one as a new row in a Google Sheet.
 *
 * SETUP
 * 1. Create a new Google Sheet. Add this header row to the first tab:
 *    Timestamp | Reference | Service | Category | Date | Time | Duration |
 *    First Name | Last Name | Phone | Email | Address | Notes
 * 2. In the Sheet: Extensions > Apps Script. Delete any starter code and
 *    paste this file's contents in.
 * 3. Click Deploy > New deployment > select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    (This is what allows the public booking page to write to the sheet
 *    without the visitor needing a Google account.)
 * 4. Copy the deployment's Web app URL (ends in /exec).
 * 5. Paste that URL into GOOGLE_SHEET_ENDPOINT near the top of the
 *    <script> block in us-stoneworks.html.
 * 6. Re-deploy (Deploy > Manage deployments > edit > new version) any time
 *    you change this script — editing the code alone does not update a
 *    live deployment.
 */

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date(),
    data.reference || '',
    data.service || '',
    data.category || '',
    data.date || '',
    data.time || '',
    data.duration || '',
    data.firstName || '',
    data.lastName || '',
    data.phone || '',
    data.email || '',
    data.address || '',
    data.notes || '',
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Lets you sanity-check the deployment by opening the /exec URL in a browser.
function doGet(e) {
  return ContentService
    .createTextOutput('U.S. Stoneworks booking endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}

<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>U.S. Stoneworks | Masonry &amp; Stone Contractors — Free Estimates</title>
<meta name="description" content="U.S. Stoneworks builds and repairs retaining walls, veneer, patios, and fireplaces for homes and businesses. Check live availability and book a free on-site estimate in minutes.">
<meta name="keywords" content="masonry contractor, stone contractor, retaining wall builder, stone veneer installation, patio masonry, fireplace mason, masonry repair, free estimate, U.S. Stoneworks">
<meta name="robots" content="index, follow">
<meta name="author" content="U.S. Stoneworks">
<link rel="canonical" href="https://us-stoneworks.github.io/us-stoneworks/">
<meta property="og:type" content="website">
<meta property="og:site_name" content="U.S. Stoneworks">
<meta property="og:title" content="U.S. Stoneworks | Masonry &amp; Stone Contractors — Free Estimates">
<meta property="og:description" content="Residential and commercial masonry — retaining walls, veneer, patios, fireplaces, and repair. Check availability and book a free estimate online.">
<meta property="og:url" content="https://us-stoneworks.github.io/us-stoneworks/">
<meta property="og:image" content="https://us-stoneworks.github.io/us-stoneworks/og-image.jpg">
<meta property="og:locale" content="en_US">
<meta name="twitter:card" content="summary_large_image">
<meta name="twitter:title" content="U.S. Stoneworks | Masonry &amp; Stone Contractors">
<meta name="twitter:description" content="Check availability and book a free on-site masonry estimate — residential and commercial.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Zilla+Slab:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@type": "GeneralContractor",
  "name": "U.S. Stoneworks",
  "description": "Residential and commercial masonry contractor specializing in retaining walls, stone veneer, patios, fireplaces, and masonry repair.",
  "url": "https://us-stoneworks.github.io/us-stoneworks",
  "telephone": "+1-555-010-2847",
  "email": "office@usstoneworks.example",
  "priceRange": "$$",
  "areaServed": "Greater metro area",
  "address": { "@type": "PostalAddress", "addressLocality": "Your City", "addressRegion": "Your State", "addressCountry": "US" },
  "openingHoursSpecification": [
    { "@type": "OpeningHoursSpecification", "dayOfWeek": ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"], "opens": "08:00", "closes": "16:00" }
  ],
  "sameAs": ["https://www.facebook.com/", "https://www.instagram.com/"],
  "makesOffer": [
    { "@type": "Offer", "name": "Residential New Project Estimate", "price": "0", "priceCurrency": "USD" },
    { "@type": "Offer", "name": "Residential Repair Estimate", "price": "0", "priceCurrency": "USD" },
    { "@type": "Offer", "name": "Commercial New Project Estimate", "price": "0", "priceCurrency": "USD" },
    { "@type": "Offer", "name": "Commercial Repair Estimate", "price": "0", "priceCurrency": "USD" }
  ]
}
</script>
<style>
  :root{
    --basalt:#211f1c;
    --slate:#2c2925;
    --slate-2:#38342e;
    --limestone:#ece5d6;
    --limestone-dim:#c9c0ac;
    --mortar:#756c5e;
    --mortar-line:#4a453d;
    --oxide:#b25330;
    --oxide-dim:#8f4126;
    --verdigris:#7c8a67;
    --focus:#e0b25a;
    --radius-s: 3px;
    --radius-m: 6px;
    font-size:16px;
  }
  *{box-sizing:border-box;}
  html{scroll-behavior:smooth;}
  body{
    margin:0;
    background:var(--basalt);
    color:var(--limestone);
    font-family:'IBM Plex Sans', sans-serif;
    line-height:1.55;
    -webkit-font-smoothing:antialiased;
  }
  h1,h2,h3,h4{
    font-family:'Zilla Slab', serif;
    font-weight:600;
    margin:0 0 .4em 0;
    letter-spacing:.2px;
  }
  a{color:inherit;}
  p{margin:0 0 1em 0;}
  .wrap{max-width:1120px;margin:0 auto;padding:0 28px;}
  ::selection{background:var(--oxide);color:var(--limestone);}
  :focus-visible{outline:2px solid var(--focus);outline-offset:3px;}
  button{font-family:inherit;cursor:pointer;}
  .mono{font-family:'IBM Plex Mono', monospace; letter-spacing:.02em;}
  .eyebrow{
    font-family:'IBM Plex Mono', monospace;
    text-transform:uppercase;
    letter-spacing:.16em;
    font-size:.72rem;
    color:var(--oxide);
    display:flex;align-items:center;gap:10px;
    margin-bottom:14px;
  }
  .eyebrow::before{content:"";width:22px;height:1px;background:var(--oxide);display:inline-block;}

  /* ---------- Coursed-stone divider (signature element) ---------- */
  .stone-course{ display:flex; width:100%; overflow:hidden; line-height:0; }
  .stone-course .row{ display:flex; width:100%; }
  .stone-course span{
    flex:0 0 auto;
    height:14px;
    border-right:2px solid var(--basalt);
    background:var(--slate-2);
  }
  .stone-course.on-slate span{ background:var(--slate); border-right-color:var(--basalt);}
  .stone-course .row:nth-child(1){ margin-left:-38px; }
  .stone-course span:nth-child(6n+1){width:9%;background:var(--mortar-line);}
  .stone-course span:nth-child(6n+2){width:6%;}
  .stone-course span:nth-child(6n+3){width:14%;background:var(--slate);}
  .stone-course span:nth-child(6n+4){width:8%;}
  .stone-course span:nth-child(6n+5){width:11%;background:var(--mortar-line);}
  .stone-course span:nth-child(6n){width:7%;}

  /* ---------- Nav ---------- */
  header{
    position:sticky;top:0;z-index:50;
    background:rgba(33,31,28,.92);
    backdrop-filter:blur(6px);
    border-bottom:1px solid var(--mortar-line);
  }
  nav{display:flex;align-items:center;justify-content:space-between;padding:16px 0;}
  .brand{display:flex;align-items:center;gap:10px;font-family:'Zilla Slab',serif;font-weight:700;font-size:1.15rem;letter-spacing:.02em;}
  .brand .mark{width:26px;height:26px;flex:none;}
  .brand small{display:block;font-family:'IBM Plex Mono',monospace;font-weight:400;font-size:.58rem;letter-spacing:.18em;color:var(--oxide);margin-top:2px;}
  .navlinks{display:flex;gap:30px;font-size:.92rem;list-style:none;margin:0;padding:0;}
  .navlinks a{text-decoration:none;color:var(--limestone-dim);transition:color .15s;}
  .navlinks a:hover{color:var(--limestone);}
  .navcta{display:flex;align-items:center;gap:18px;}
  .navcta .phone{font-family:'IBM Plex Mono',monospace;font-size:.85rem;color:var(--limestone-dim);display:none;}
  .btn{
    display:inline-flex;align-items:center;gap:8px;
    padding:11px 20px;border-radius:var(--radius-s);
    border:1px solid transparent;
    font-size:.88rem;font-weight:500;text-decoration:none;
    transition:transform .15s ease, background .15s, border-color .15s;
  }
  .btn-primary{background:var(--oxide);color:var(--limestone);}
  .btn-primary:hover{background:var(--oxide-dim); transform:translateY(-1px);}
  .btn-ghost{background:transparent;border-color:var(--mortar-line);color:var(--limestone);}
  .btn-ghost:hover{border-color:var(--limestone-dim);}
  .btn:disabled{opacity:.35;cursor:not-allowed;transform:none;}
  .mobile-toggle{display:none;}

  /* ---------- Hero ---------- */
  .hero{
    padding:88px 0 0 0;
    position:relative;
  }
  .hero-grid{
    display:grid;grid-template-columns:1.1fr .8fr;gap:60px;align-items:center;
    padding-bottom:70px;
  }
  .hero h1{font-size:clamp(2.4rem, 5vw, 3.6rem);line-height:1.04;color:var(--limestone);}
  .hero h1 em{font-style:normal;color:var(--oxide);}
  .hero p.lead{font-size:1.08rem;color:var(--limestone-dim);max-width:46ch;}
  .hero-ctas{display:flex;gap:14px;margin-top:28px;flex-wrap:wrap;}
  .stat-row{display:flex;gap:32px;margin-top:46px;flex-wrap:wrap;}
  .stat{font-family:'IBM Plex Mono',monospace;}
  .stat b{display:block;font-family:'Zilla Slab',serif;font-size:1.5rem;color:var(--limestone);font-weight:600;}
  .stat span{font-size:.72rem;color:var(--mortar);text-transform:uppercase;letter-spacing:.1em;}

  .wall{
    aspect-ratio:1/1.05;
    border-radius:var(--radius-m);
    position:relative;
    overflow:hidden;
    background:var(--slate);
    border:1px solid var(--mortar-line);
    display:grid;
    grid-template-columns:repeat(5,1fr);
    grid-auto-rows:22px;
    gap:3px;
    padding:14px;
  }
  .wall div{background:var(--slate-2);border-radius:2px;}
  .wall div:nth-child(4n+1){background:#332f29;}
  .wall div:nth-child(7n+2){background:#403b33;}
  .wall div:nth-child(11n+3){background:var(--oxide-dim);opacity:.55;}
  .wall div:nth-child(9n+4){grid-column:span 2;}
  .wall div:nth-child(13n+5){grid-column:span 2;}

  /* ---------- Section basics ---------- */
  section{padding:76px 0;}
  .section-head{max-width:60ch;margin-bottom:44px;}
  .section-head h2{font-size:clamp(1.7rem,3vw,2.3rem);}
  .section-head p{color:var(--limestone-dim);}

  /* ---------- Services ---------- */
  .service-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;}
  .service-card{
    background:var(--slate);
    border:1px solid var(--mortar-line);
    border-radius:var(--radius-m);
    padding:24px;
    display:flex;flex-direction:column;gap:12px;
    transition:border-color .15s, transform .15s;
  }
  .service-card:hover{border-color:var(--oxide-dim);transform:translateY(-3px);}
  .service-card .tag{font-family:'IBM Plex Mono',monospace;font-size:.68rem;letter-spacing:.1em;text-transform:uppercase;color:var(--verdigris);}
  .service-card h3{font-size:1.15rem;color:var(--limestone);}
  .service-card p{font-size:.9rem;color:var(--limestone-dim);flex:1;margin:0;}
  .service-meta{display:flex;justify-content:space-between;align-items:center;font-family:'IBM Plex Mono',monospace;font-size:.78rem;color:var(--mortar);border-top:1px solid var(--mortar-line);padding-top:12px;}
  .service-meta .price{color:var(--verdigris);}
  .service-card button{
    margin-top:6px;width:100%;padding:10px;border-radius:var(--radius-s);
    background:transparent;border:1px solid var(--mortar-line);color:var(--limestone);font-size:.85rem;
    transition:background .15s,border-color .15s;
  }
  .service-card button:hover{background:var(--oxide);border-color:var(--oxide);}

  /* ---------- Process ---------- */
  .process{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;}
  .process .step{padding:22px 0;border-top:1px solid var(--mortar-line);}
  .process .step .num{font-family:'IBM Plex Mono',monospace;color:var(--oxide);font-size:.85rem;}
  .process .step h4{margin-top:10px;font-size:1.05rem;}
  .process .step p{font-size:.88rem;color:var(--limestone-dim);margin:0;}

  /* ---------- Booking ---------- */
  .booking{background:var(--slate);border-top:1px solid var(--mortar-line);border-bottom:1px solid var(--mortar-line);}
  .progress{display:flex;gap:8px;margin-bottom:36px;max-width:520px;}
  .progress .seg{flex:1;height:8px;border-radius:2px;background:var(--slate-2);position:relative;overflow:hidden;}
  .progress .seg.done{background:var(--oxide);}
  .progress .seg.active{background:linear-gradient(90deg,var(--oxide) 50%, var(--slate-2) 50%);}
  .progress-labels{display:flex;justify-content:space-between;max-width:520px;font-family:'IBM Plex Mono',monospace;font-size:.66rem;text-transform:uppercase;letter-spacing:.08em;color:var(--mortar);margin-bottom:40px;}
  .progress-labels span.on{color:var(--oxide);}

  .panel{display:none;}
  .panel.active{display:block;animation:rise .35s ease;}
  @keyframes rise{from{opacity:0;transform:translateY(8px);}to{opacity:1;transform:translateY(0);}}

  .pick-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px;}
  .pick-card{
    text-align:left;background:var(--basalt);border:1px solid var(--mortar-line);border-radius:var(--radius-m);
    padding:20px;display:flex;flex-direction:column;gap:6px;transition:border-color .15s, background .15s;
  }
  .pick-card:hover{border-color:var(--verdigris);}
  .pick-card.selected{border-color:var(--oxide);background:#2a201a;}
  .pick-card h3{font-size:1.02rem;color:var(--limestone);}
  .pick-card span.tag{font-family:'IBM Plex Mono',monospace;font-size:.65rem;color:var(--verdigris);text-transform:uppercase;letter-spacing:.08em;}
  .pick-card p{font-size:.83rem;color:var(--limestone-dim);margin:0;}

  .cal-wrap{display:grid;grid-template-columns:1.15fr 1fr;gap:40px;}
  .cal-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;}
  .cal-head h3{font-size:1.1rem;margin:0;}
  .cal-nav{display:flex;gap:6px;}
  .cal-nav button{background:var(--basalt);border:1px solid var(--mortar-line);color:var(--limestone);width:32px;height:32px;border-radius:var(--radius-s);}
  .cal-nav button:hover{border-color:var(--verdigris);}
  .dow{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;font-family:'IBM Plex Mono',monospace;font-size:.62rem;text-transform:uppercase;color:var(--mortar);margin-bottom:6px;text-align:center;}
  .cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;}
  .day{
    aspect-ratio:1;display:flex;align-items:center;justify-content:center;
    font-family:'IBM Plex Mono',monospace;font-size:.82rem;
    background:var(--basalt);color:var(--limestone-dim);
    border:1px solid var(--mortar-line);
    cursor:pointer; position:relative;
  }
  .day:nth-child(5n+1){border-radius:5px 2px 3px 6px;}
  .day:nth-child(5n+2){border-radius:2px 6px 2px 4px;}
  .day:nth-child(5n+3){border-radius:6px 3px 5px 2px;}
  .day:nth-child(5n+4){border-radius:3px 2px 6px 4px;}
  .day:nth-child(5n){border-radius:2px 4px 2px 5px;}
  .day.empty{background:transparent;border-color:transparent;cursor:default;}
  .day.closed{color:var(--mortar);cursor:not-allowed;background:var(--slate-2);}
  .day.past{color:var(--mortar);cursor:not-allowed;opacity:.35;}
  .day.open:hover{border-color:var(--verdigris);transform:translateY(-2px);}
  .day.today::after{content:"";position:absolute;bottom:4px;width:4px;height:4px;border-radius:50%;background:var(--oxide);}
  .day.selected{background:var(--oxide);border-color:var(--oxide);color:var(--limestone);box-shadow:inset 0 2px 6px rgba(0,0,0,.35);}

  .slots-col h3{font-size:1.1rem;margin-bottom:16px;}
  .slot-list{display:flex;flex-direction:column;gap:8px;max-height:340px;overflow:auto;padding-right:4px;}
  .slot{
    display:flex;justify-content:space-between;align-items:center;
    padding:11px 14px;border:1px solid var(--mortar-line);border-radius:var(--radius-s);
    background:var(--basalt);font-family:'IBM Plex Mono',monospace;font-size:.85rem;
    transition:border-color .15s, background .15s;
  }
  .slot.available:hover{border-color:var(--verdigris);cursor:pointer;}
  .slot.selected{background:var(--oxide);border-color:var(--oxide);}
  .slot.booked{opacity:.35;text-decoration:line-through;cursor:not-allowed;}
  .slot .status{font-size:.65rem;text-transform:uppercase;letter-spacing:.06em;color:var(--verdigris);}
  .slot.booked .status{color:var(--mortar);}
  .empty-hint{color:var(--mortar);font-size:.88rem;font-style:italic;}

  form.contact-form{display:grid;grid-template-columns:1fr 1fr;gap:16px;max-width:680px;}
  form.contact-form .full{grid-column:1/-1;}
  label{display:block;font-family:'IBM Plex Mono',monospace;font-size:.68rem;text-transform:uppercase;letter-spacing:.08em;color:var(--mortar);margin-bottom:6px;}
  input,textarea{
    width:100%;padding:11px 12px;background:var(--basalt);border:1px solid var(--mortar-line);
    border-radius:var(--radius-s);color:var(--limestone);font-family:'IBM Plex Sans',sans-serif;font-size:.92rem;
  }
  input:focus,textarea:focus{border-color:var(--oxide);}
  textarea{resize:vertical;min-height:80px;}

  .summary-card{
    background:var(--basalt);border:1px solid var(--mortar-line);border-radius:var(--radius-m);
    padding:22px 26px;margin-bottom:24px;max-width:520px;
  }
  .summary-card .row{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--mortar-line);font-size:.9rem;}
  .summary-card .row:last-child{border-bottom:none;}
  .summary-card .row span:first-child{color:var(--mortar);font-family:'IBM Plex Mono',monospace;font-size:.72rem;text-transform:uppercase;letter-spacing:.06em;}

  .confirm-box{text-align:left;max-width:560px;}
  .confirm-box .ref{
    display:inline-block;font-family:'IBM Plex Mono',monospace;color:var(--oxide);
    border:1px dashed var(--oxide-dim);padding:8px 14px;border-radius:var(--radius-s);margin-bottom:20px;font-size:.85rem;
  }
  .step-actions{display:flex;justify-content:space-between;margin-top:32px;max-width:680px;}

  footer{padding:56px 0 34px 0;}
  .foot-grid{display:grid;grid-template-columns:1.4fr 1fr 1fr 1fr;gap:40px;margin-bottom:40px;}
  footer h4{font-size:.95rem;color:var(--limestone);}
  footer .brand{margin-bottom:12px;}
  footer p{color:var(--limestone-dim);font-size:.88rem;}
  .foot-links{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:8px;font-size:.88rem;}
  .foot-links a{text-decoration:none;color:var(--limestone-dim);}
  .foot-links a:hover{color:var(--limestone);}
  .foot-bottom{display:flex;justify-content:space-between;align-items:center;border-top:1px solid var(--mortar-line);padding-top:22px;font-size:.78rem;color:var(--mortar);flex-wrap:wrap;gap:10px;}
  .social{display:flex;gap:10px;}
  .social a{width:32px;height:32px;border:1px solid var(--mortar-line);border-radius:50%;display:flex;align-items:center;justify-content:center;text-decoration:none;}
  .social a:hover{border-color:var(--oxide);}
  .placeholder-note{font-size:.7rem;color:var(--mortar);margin-top:10px;}

  @media (max-width:880px){
    .navlinks{display:none;}
    .hero-grid{grid-template-columns:1fr;}
    .service-grid{grid-template-columns:repeat(2,1fr);}
    .process{grid-template-columns:repeat(2,1fr);}
    .cal-wrap{grid-template-columns:1fr;}
    .foot-grid{grid-template-columns:1fr 1fr;}
    form.contact-form{grid-template-columns:1fr;}
  }
  @media (max-width:520px){
    .service-grid{grid-template-columns:1fr;}
    .process{grid-template-columns:1fr;}
    .foot-grid{grid-template-columns:1fr;}
    .stat-row{gap:20px;}
  }
  @media (prefers-reduced-motion: reduce){
    html{scroll-behavior:auto;}
    *{animation-duration:.01ms !important;transition-duration:.01ms !important;}
  }
</style>
</head>
<body>

<header>
  <div class="wrap">
    <nav>
      <div class="brand">
        <svg class="mark" viewBox="0 0 26 26" fill="none" xmlns="http://www.w3.org/2000/svg">
          <rect x="1" y="15" width="8" height="9" rx="1" fill="#b25330"/>
          <rect x="10" y="15" width="6" height="9" rx="1" fill="#ece5d6"/>
          <rect x="17" y="15" width="8" height="9" rx="1" fill="#7c8a67"/>
          <rect x="4" y="7" width="9" height="7" rx="1" fill="#ece5d6"/>
          <rect x="14" y="7" width="8" height="7" rx="1" fill="#b25330" opacity=".8"/>
          <rect x="8" y="1" width="10" height="5" rx="1" fill="#ece5d6" opacity=".6"/>
        </svg>
        <div>U.S. STONEWORKS<small>MASONRY &amp; STONE CONTRACTORS</small></div>
      </div>
      <ul class="navlinks">
        <li><a href="#services">Services</a></li>
        <li><a href="#process">Process</a></li>
        <li><a href="#booking">Availability</a></li>
      </ul>
      <div class="navcta">
        <span class="phone mono">(555) 010-2847</span>
        <a class="btn btn-primary" href="#booking">Check Availability</a>
      </div>
    </nav>
  </div>
</header>

<main>
<section class="hero" aria-label="Introduction">
  <div class="wrap">
    <div class="hero-grid">
      <div>
        <div class="eyebrow">Free On-Site Estimates</div>
        <h1>Built <em>stone</em><br>by stone.</h1>
        <p class="lead">U.S. Stoneworks handles residential and commercial masonry — retaining walls, veneer, patios, fireplaces, and structural repair. See what's open on our crew's calendar and lock in a walkthrough.</p>
        <div class="hero-ctas">
          <a class="btn btn-primary" href="#booking">Check Availability</a>
          <a class="btn btn-ghost" href="#services">View Services</a>
        </div>
        <div class="stat-row">
          <div class="stat"><b>22 yrs</b><span>In the trade</span></div>
          <div class="stat"><b>Lic. &amp; Bonded</b><span>State contractor</span></div>
          <div class="stat"><b>1 hr</b><span>Typical estimate visit</span></div>
        </div>
      </div>
      <div class="wall" aria-hidden="true">
        <div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div>
        <div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div>
        <div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div>
        <div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div>
      </div>
    </div>
  </div>
  <div class="stone-course" aria-hidden="true">
    <div class="row">
      <span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span>
    </div>
  </div>
</section>

<section id="services">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">Estimate Types</div>
      <h2>Pick the estimate that fits your project</h2>
      <p>Every visit is free and gets you a written quote — no obligation. Select a service below to jump straight into booking.</p>
    </div>
    <div class="service-grid" id="serviceGridTop"></div>
  </div>
</section>

<div class="stone-course on-slate" aria-hidden="true">
  <div class="row">
    <span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span><span></span>
  </div>
</div>

<section id="process">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">How It Works</div>
      <h2>From first call to finished wall</h2>
    </div>
    <div class="process">
      <div class="step"><div class="num">01</div><h4>Pick a project type</h4><p>Tell us if it's new construction or a repair, residential or commercial.</p></div>
      <div class="step"><div class="num">02</div><h4>Choose a time</h4><p>Grab an open slot on the crew's calendar — mornings and afternoons available.</p></div>
      <div class="step"><div class="num">03</div><h4>We walk the site</h4><p>A lead mason measures, checks access, and talks through materials on-site.</p></div>
      <div class="step"><div class="num">04</div><h4>Get your quote</h4><p>A written estimate follows within 48 hours, no pressure to commit.</p></div>
    </div>
  </div>
</section>

<section class="booking" id="booking">
  <div class="wrap">
    <div class="section-head">
      <div class="eyebrow">Availability</div>
      <h2>Schedule your free estimate</h2>
      <p>Four short steps. You'll get a confirmation with your reference number at the end.</p>
    </div>

    <div class="progress">
      <div class="seg" data-seg="1"></div>
      <div class="seg" data-seg="2"></div>
      <div class="seg" data-seg="3"></div>
      <div class="seg" data-seg="4"></div>
    </div>
    <div class="progress-labels">
      <span data-label="1">Service</span>
      <span data-label="2">Date &amp; Time</span>
      <span data-label="3">Your Details</span>
      <span data-label="4">Confirmed</span>
    </div>

    <!-- STEP 1 -->
    <div class="panel active" id="panel-1">
      <div class="pick-grid" id="serviceGrid"></div>
      <div class="step-actions">
        <span></span>
        <button class="btn btn-primary" id="toStep2" disabled>Continue to Date &amp; Time</button>
      </div>
    </div>

    <!-- STEP 2 -->
    <div class="panel" id="panel-2">
      <div class="cal-wrap">
        <div>
          <div class="cal-head">
            <h3 id="calMonthLabel"></h3>
            <div class="cal-nav">
              <button id="prevMonth" aria-label="Previous month">&lsaquo;</button>
              <button id="nextMonth" aria-label="Next month">&rsaquo;</button>
            </div>
          </div>
          <div class="dow"><span>S</span><span>M</span><span>T</span><span>W</span><span>T</span><span>F</span><span>S</span></div>
          <div class="cal-grid" id="calGrid"></div>
          <p class="placeholder-note">Closed Sundays · Booking closes at 3:00 PM for a 1-hour visit</p>
        </div>
        <div class="slots-col">
          <h3 id="slotsHeading">Select a date</h3>
          <div class="slot-list" id="slotList"><p class="empty-hint">Available times appear here once you choose a day.</p></div>
        </div>
      </div>
      <div class="step-actions">
        <button class="btn btn-ghost" data-back="1">Back</button>
        <button class="btn btn-primary" id="toStep3" disabled>Continue to Your Details</button>
      </div>
    </div>

    <!-- STEP 3 -->
    <div class="panel" id="panel-3">
      <form class="contact-form" id="contactForm">
        <div>
          <label for="fname">First Name</label>
          <input type="text" id="fname" required>
        </div>
        <div>
          <label for="lname">Last Name</label>
          <input type="text" id="lname" required>
        </div>
        <div>
          <label for="phone">Phone</label>
          <input type="tel" id="phone" required>
        </div>
        <div>
          <label for="email">Email</label>
          <input type="email" id="email" required>
        </div>
        <div class="full">
          <label for="address">Project Address</label>
          <input type="text" id="address" placeholder="Street, City, State" required>
        </div>
        <div class="full">
          <label for="notes">Project Notes (optional)</label>
          <textarea id="notes" placeholder="Wall length, stone type, gate access, timeline — anything that helps us prepare."></textarea>
        </div>
      </form>
      <div class="step-actions">
        <button class="btn btn-ghost" data-back="2">Back</button>
        <button class="btn btn-primary" id="toStep4">Confirm Booking</button>
      </div>
    </div>

    <!-- STEP 4 -->
    <div class="panel" id="panel-4">
      <div class="confirm-box">
        <div class="eyebrow">Booking Requested</div>
        <h2>You're on the calendar.</h2>
        <p style="color:var(--limestone-dim)">A lead mason will confirm by phone or email before the visit. Save your reference number below.</p>
        <div class="ref mono" id="refNumber">REF-000000</div>
        <div class="summary-card" id="summaryCard"></div>
        <a class="btn btn-ghost" href="#services" id="bookAnother">Book Another Estimate</a>
      </div>
    </div>
  </div>
</section>
</main>

<footer>
  <div class="wrap">
    <div class="foot-grid">
      <div>
        <div class="brand">
          <svg class="mark" viewBox="0 0 26 26" fill="none" xmlns="http://www.w3.org/2000/svg">
            <rect x="1" y="15" width="8" height="9" rx="1" fill="#b25330"/>
            <rect x="10" y="15" width="6" height="9" rx="1" fill="#ece5d6"/>
            <rect x="17" y="15" width="8" height="9" rx="1" fill="#7c8a67"/>
            <rect x="4" y="7" width="9" height="7" rx="1" fill="#ece5d6"/>
            <rect x="14" y="7" width="8" height="7" rx="1" fill="#b25330" opacity=".8"/>
            <rect x="8" y="1" width="10" height="5" rx="1" fill="#ece5d6" opacity=".6"/>
          </svg>
          <div>U.S. STONEWORKS</div>
        </div>
        <p>Residential and commercial masonry contractors — retaining walls, veneer, patios, fireplaces, and repair.</p>
        <div class="social">
          <a href="#" aria-label="Facebook">f</a>
          <a href="#" aria-label="Instagram">ig</a>
        </div>
      </div>
      <div>
        <h4>Company</h4>
        <ul class="foot-links">
          <li><a href="#services">Services</a></li>
          <li><a href="#process">Our Process</a></li>
          <li><a href="#booking">Availability</a></li>
        </ul>
      </div>
      <div>
        <h4>Hours</h4>
        <ul class="foot-links">
          <li class="mono">Mon–Sat · 8:00–16:00</li>
          <li class="mono">Sunday · Closed</li>
        </ul>
      </div>
      <div>
        <h4>Contact</h4>
        <ul class="foot-links">
          <li class="mono">(555) 010-2847</li>
          <li class="mono">office@usstoneworks.example</li>
          <li>Serving the greater metro area</li>
        </ul>
      </div>
    </div>
    <div class="foot-bottom">
      <span>© 2026 U.S. Stoneworks. Lic. #000000 (placeholder).</span>
      <span>Site content is placeholder — swap in real copy, license number, and contact details.</span>
    </div>
  </div>
</footer>

<script>
(function(){
  const services = [
    {id:'res-new', category:'Residential', name:'New Project Estimate', duration:'1 hr', price:'Free',
      desc:'Patios, walls, veneer, or fireplaces — ground-up residential stonework.'},
    {id:'res-repair', category:'Residential', name:'Repair Estimate', duration:'45 min', price:'Free',
      desc:'Cracked mortar, settling walls, or storm damage on existing stonework.'},
    {id:'com-new', category:'Commercial', name:'New Project Estimate', duration:'1 hr', price:'Free',
      desc:'Facades, retaining structures, and hardscape for commercial sites.'},
    {id:'com-repair', category:'Commercial', name:'Repair Estimate', duration:'45 min', price:'Free',
      desc:'Structural or cosmetic repair for commercial masonry and stonework.'}
  ];

  const state = { service:null, date:null, time:null, monthOffset:0 };

  // ---------- Google Sheets logging (optional) ----------
  // Paste the /exec URL from your deployed Apps Script web app here to log
  // every booking as a new row in a Google Sheet. Leave blank to disable.
  // See google-apps-script-backend.gs (provided alongside this file) for the
  // matching server-side script and deployment steps.
  const GOOGLE_SHEET_ENDPOINT = "";

  function logBookingToSheet(payload){
    if(!GOOGLE_SHEET_ENDPOINT) return;
    fetch(GOOGLE_SHEET_ENDPOINT, {
      method: "POST",
      mode: "no-cors", // Apps Script web apps don't return CORS headers by default;
                        // no-cors lets the write through even though we can't read the response.
      headers: { "Content-Type": "text/plain;charset=utf-8" }, // avoids a CORS preflight
      body: JSON.stringify(payload)
    }).catch(()=>{ /* fails silently so a network hiccup never blocks the on-screen confirmation */ });
  }

  const serviceGridTop = document.getElementById('serviceGridTop');
  const serviceGrid = document.getElementById('serviceGrid');

  function serviceCardHTML(s, variant){
    return `
      <div class="${variant} pick-card" data-id="${s.id}" tabindex="0" role="button" aria-pressed="false">
        <span class="tag">${s.category}</span>
        <h3>${s.name}</h3>
        <p>${s.desc}</p>
        <div class="service-meta"><span>${s.duration}</span><span class="price">${s.price}</span></div>
      </div>`;
  }

  serviceGridTop.innerHTML = services.map(s => `
    <div class="service-card">
      <span class="tag">${s.category}</span>
      <h3>${s.name}</h3>
      <p>${s.desc}</p>
      <div class="service-meta"><span class="mono">${s.duration}</span><span class="price mono">${s.price}</span></div>
      <button data-jump="${s.id}">Select &amp; Check Availability</button>
    </div>`).join('');

  serviceGrid.innerHTML = services.map(s => serviceCardHTML(s,'service-card')).join('');

  function selectService(id, jump){
    state.service = services.find(s => s.id === id);
    document.querySelectorAll('#serviceGrid .pick-card').forEach(el=>{
      const on = el.dataset.id === id;
      el.classList.toggle('selected', on);
      el.setAttribute('aria-pressed', on);
    });
    document.getElementById('toStep2').disabled = false;
    if(jump) goToStep(2);
  }

  serviceGrid.addEventListener('click', e=>{
    const card = e.target.closest('.pick-card');
    if(card) selectService(card.dataset.id, false);
  });
  serviceGrid.addEventListener('keydown', e=>{
    const card = e.target.closest('.pick-card');
    if(card && (e.key==='Enter'||e.key===' ')){ e.preventDefault(); selectService(card.dataset.id,false); }
  });
  serviceGridTop.addEventListener('click', e=>{
    const btn = e.target.closest('button[data-jump]');
    if(btn){ selectService(btn.dataset.jump, true); }
  });

  // ---------- Calendar ----------
  const calGrid = document.getElementById('calGrid');
  const calMonthLabel = document.getElementById('calMonthLabel');
  const slotList = document.getElementById('slotList');
  const slotsHeading = document.getElementById('slotsHeading');
  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];

  const today = new Date();
  today.setHours(0,0,0,0);

  function hashDate(d){
    const s = d.toISOString().slice(0,10);
    let h = 0;
    for(let i=0;i<s.length;i++){ h = (h*31 + s.charCodeAt(i)) >>> 0; }
    return h;
  }

  function slotsForDate(d){
    const hours = [8,9,10,11,12,13,14,15];
    const h = hashDate(d);
    return hours.map((hr,i)=>{
      const booked = ((h >> i) & 1) === 1 && i !== 0;
      const label = (hr % 12 === 0 ? 12 : hr % 12) + ':00 ' + (hr < 12 ? 'AM' : 'PM');
      return { hour:hr, label, booked };
    });
  }

  function renderCalendar(){
    const base = new Date(today.getFullYear(), today.getMonth() + state.monthOffset, 1);
    calMonthLabel.textContent = monthNames[base.getMonth()] + ' ' + base.getFullYear();
    document.getElementById('prevMonth').disabled = state.monthOffset <= 0;

    calGrid.innerHTML = '';
    const firstDow = base.getDay();
    const daysInMonth = new Date(base.getFullYear(), base.getMonth()+1, 0).getDate();

    for(let i=0;i<firstDow;i++){
      const e = document.createElement('div'); e.className = 'day empty'; calGrid.appendChild(e);
    }
    for(let d=1; d<=daysInMonth; d++){
      const date = new Date(base.getFullYear(), base.getMonth(), d);
      const cell = document.createElement('div');
      cell.className = 'day';
      cell.textContent = d;
      const isPast = date < today;
      const isSunday = date.getDay() === 0;
      const isToday = date.getTime() === today.getTime();
      if(isToday) cell.classList.add('today');
      if(isPast){ cell.classList.add('past'); }
      else if(isSunday){ cell.classList.add('closed'); }
      else {
        cell.classList.add('open');
        cell.addEventListener('click', ()=> selectDate(date, cell));
      }
      if(state.date && date.getTime() === state.date.getTime()) cell.classList.add('selected');
      calGrid.appendChild(cell);
    }
  }

  function selectDate(date, cell){
    state.date = date;
    state.time = null;
    document.querySelectorAll('.day.selected').forEach(el=>el.classList.remove('selected'));
    cell.classList.add('selected');
    const label = date.toLocaleDateString('en-US', { weekday:'long', month:'long', day:'numeric' });
    slotsHeading.textContent = label;
    const slots = slotsForDate(date);
    slotList.innerHTML = slots.map(s => `
      <div class="slot ${s.booked ? 'booked' : 'available'}" data-hour="${s.hour}">
        <span class="mono">${s.label}</span>
        <span class="status">${s.booked ? 'Booked' : 'Open'}</span>
      </div>`).join('');
    document.getElementById('toStep3').disabled = true;
  }

  slotList.addEventListener('click', e=>{
    const slot = e.target.closest('.slot.available');
    if(!slot) return;
    document.querySelectorAll('.slot.selected').forEach(el=>el.classList.remove('selected'));
    slot.classList.add('selected');
    state.time = slot.querySelector('.mono').textContent;
    document.getElementById('toStep3').disabled = false;
  });

  document.getElementById('prevMonth').addEventListener('click', ()=>{ if(state.monthOffset>0){ state.monthOffset--; renderCalendar(); }});
  document.getElementById('nextMonth').addEventListener('click', ()=>{ if(state.monthOffset<2){ state.monthOffset++; renderCalendar(); }});

  // ---------- Step navigation ----------
  function goToStep(n){
    document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
    document.getElementById('panel-'+n).classList.add('active');
    document.querySelectorAll('.progress .seg').forEach(seg=>{
      const i = Number(seg.dataset.seg);
      seg.classList.toggle('done', i < n);
      seg.classList.toggle('active', i === n);
    });
    document.querySelectorAll('.progress-labels span').forEach(l=>{
      l.classList.toggle('on', Number(l.dataset.label) <= n);
    });
    document.getElementById('booking').scrollIntoView({behavior:'smooth', block:'start'});
    if(n===2) renderCalendar();
  }

  document.getElementById('toStep2').addEventListener('click', ()=> goToStep(2));
  document.getElementById('toStep3').addEventListener('click', ()=> goToStep(3));
  document.querySelectorAll('[data-back]').forEach(btn=>{
    btn.addEventListener('click', ()=> goToStep(Number(btn.dataset.back)));
  });

  document.getElementById('toStep4').addEventListener('click', ()=>{
    const form = document.getElementById('contactForm');
    if(!form.reportValidity()) return;
    const fname = document.getElementById('fname').value.trim();
    const lname = document.getElementById('lname').value.trim();
    const phone = document.getElementById('phone').value.trim();
    const email = document.getElementById('email').value.trim();
    const address = document.getElementById('address').value.trim();

    const ref = 'USW-' + Math.floor(100000 + Math.random()*899999);
    document.getElementById('refNumber').textContent = ref;

    const dateLabel = state.date.toLocaleDateString('en-US', { weekday:'long', month:'long', day:'numeric', year:'numeric' });
    document.getElementById('summaryCard').innerHTML = `
      <div class="row"><span>Service</span><span>${state.service.name} (${state.service.category})</span></div>
      <div class="row"><span>Date</span><span>${dateLabel}</span></div>
      <div class="row"><span>Time</span><span>${state.time}</span></div>
      <div class="row"><span>Duration</span><span>${state.service.duration}</span></div>
      <div class="row"><span>Client</span><span>${fname} ${lname}</span></div>
      <div class="row"><span>Phone</span><span>${phone}</span></div>
      <div class="row"><span>Email</span><span>${email}</span></div>
      <div class="row"><span>Site Address</span><span>${address}</span></div>
    `;

    logBookingToSheet({
      reference: ref,
      service: state.service.name,
      category: state.service.category,
      date: dateLabel,
      time: state.time,
      duration: state.service.duration,
      firstName: fname,
      lastName: lname,
      phone: phone,
      email: email,
      address: address,
      notes: document.getElementById('notes').value.trim(),
      submittedAt: new Date().toISOString()
    });

    goToStep(4);
  });

  document.getElementById('bookAnother').addEventListener('click', (e)=>{
    e.preventDefault();
    state.service=null; state.date=null; state.time=null; state.monthOffset=0;
    document.getElementById('contactForm').reset();
    document.querySelectorAll('#serviceGrid .pick-card').forEach(el=>el.classList.remove('selected'));
    document.getElementById('toStep2').disabled = true;
    document.getElementById('toStep3').disabled = true;
    slotList.innerHTML = '<p class="empty-hint">Available times appear here once you choose a day.</p>';
    slotsHeading.textContent = 'Select a date';
    goToStep(1);
  });

  renderCalendar();
})();
</script>
</body>
</html>

# U.S. Stoneworks — Website & Booking Page

A single-page site for U.S. Stoneworks with an interactive, step-by-step
availability/booking widget (service → date & time → contact details →
confirmation), SEO metadata, and optional logging of bookings to a Google
Sheet via Apps Script.

## Files

- `index.html` — the entire site (HTML/CSS/JS, one file, no build step)
- `google-apps-script-backend.gs` — optional server-side script that logs
  each booking into a Google Sheet

## Deploy with GitHub Pages (free)

Repo: https://github.com/US-Stoneworks/us-stoneworks

1. **Repo already created** ✅ — `US-Stoneworks/us-stoneworks` on GitHub.
2. **Upload these files** via *Add file → Upload files* on the repo page:
   `index.html`, `google-apps-script-backend.gs`, `README.md`. Commit.
3. **Turn on Pages.** In the repo: *Settings → Pages → Build and
   deployment → Source: Deploy from a branch → Branch: `main`, folder
   `/ (root)` → Save.
4. GitHub will publish the site at:
   `https://us-stoneworks.github.io/us-stoneworks/`
   (takes 1–2 minutes on first deploy; every future commit to `main`
   redeploys it automatically.)
5. The SEO placeholder URLs (canonical, Open Graph, JSON-LD) in
   `index.html` are already set to this address — no find/replace needed.
   Still update the phone/email/address placeholders in the footer and
   JSON-LD block with the real business details.
6. **Custom domain (optional).** *Settings → Pages → Custom domain*, add
   your domain, then create a `CNAME` record at your DNS provider pointing
   to `us-stoneworks.github.io`. GitHub issues HTTPS automatically once DNS
   resolves.

## Wiring up booking logs to Google Sheets (optional)

See the setup comment at the top of `google-apps-script-backend.gs`. Once
deployed, paste the Apps Script Web App `/exec` URL into the
`GOOGLE_SHEET_ENDPOINT` constant near the top of the `<script>` block in
`index.html`, commit, and push — bookings will start appending to the
sheet as rows.

## Notes

- No build tools or dependencies — it's plain HTML/CSS/JS, so GitHub Pages
  can serve it directly from the repo root.
- All placeholder content (phone, email, license #, social links, address)
  is marked in the footer and should be swapped for the real business
  details before going live.
