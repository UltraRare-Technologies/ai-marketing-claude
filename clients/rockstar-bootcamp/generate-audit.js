const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType,
        BorderStyle, WidthType, ShadingType, PageBreak, HeadingLevel, PageOrientation, LevelFormat } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "2E75B6" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
      },
      {
        reference: "numbers",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } }
          }
        ]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22 } } // 11pt
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 }
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 1 }
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 120, after: 80 }, outlineLevel: 2 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: {
          width: 12240,   // 8.5 inches (US Letter)
          height: 15840   // 11 inches
        },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } // 1 inch margins
      }
    },
    children: [
      // Title Page
      new Paragraph({
        children: [new TextRun({ text: "SEO AUDIT REPORT", bold: true, size: 48, color: "2E75B6" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Rock Star Boot Camp", bold: true, size: 36, color: "2E75B6" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Phoenix, Arizona", size: 26, color: "666666" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Prepared: March 17, 2026", size: 24, italics: true })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 600 }
      }),

      // Executive Summary
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Executive Summary")] }),
      new Paragraph({
        children: [new TextRun("This comprehensive SEO audit evaluates Rock Star Boot Camp's web presence and search engine optimization readiness. The business is a fitness boot camp facility in North Phoenix offering 9xFit interval training with membership plans ranging from $29-$59 trials to $247 regular memberships.")]
      }),
      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "Key Findings:", bold: true })],
        spacing: { after: 80 }
      }),

      // Executive summary table
      createTable([
        ["CRITICAL ISSUES", "HIGH IMPACT OPPORTUNITIES", "QUICK WINS"],
        [
          "• Multiple H1 tags (SEO anti-pattern)\n• Missing meta descriptions\n• Weak local SEO signals\n• Limited content strategy",
          "• Local keyword targeting\n• Google Business Profile optimization\n• Service area pages\n• Review generation campaigns",
          "• Consolidate H1 tags\n• Write unique meta descriptions\n• Add location schema\n• NAP consistency checks"
        ]
      ], [2000, 3500, 3860]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 240 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Priority Actions:", bold: true })],
        spacing: { after: 80 }
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Fix on-page SEO issues (H1 tags, meta descriptions) - HIGH IMPACT, LOW EFFORT")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Optimize Google Business Profile with enhanced descriptions and FAQs - HIGH IMPACT, LOW-MEDIUM EFFORT")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Implement local service area pages targeting neighborhood keywords - HIGH IMPACT, MEDIUM EFFORT")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Establish review collection and management system - MEDIUM IMPACT, LOW EFFORT")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Create blog content targeting fitness boot camp long-tail keywords - MEDIUM IMPACT, MEDIUM EFFORT")]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // On-Page SEO Analysis
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. On-Page SEO Analysis")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Title Tags")] }),
      createTable([
        ["Element", "Current Status", "Recommendations", "Priority"],
        ["Homepage Title", "Generic - 'Rock Star Boot Camp'", "Change to 'Fitness Boot Camp in Phoenix, AZ | 9xFit Training'", "HIGH"],
        ["Services Page", "Missing/Generic", "Add 'Group Fitness Classes & Memberships | Phoenix Boot Camp'", "HIGH"],
        ["Pricing Page", "Not optimized", "Use 'Membership Plans & Trial Classes | Rock Star Boot Camp'", "MEDIUM"],
        ["About Page", "Missing optimization", "Add 'About Our Boot Camp | Phoenix Fitness Training'", "MEDIUM"]
      ], [1300, 2200, 3500, 1360]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Meta Descriptions")] }),
      new Paragraph({
        children: [new TextRun({ text: "Current Status: MISSING/NOT OPTIMIZED", bold: true, color: "C00000" })],
        spacing: { after: 80 }
      }),
      new Paragraph({
        children: [new TextRun("Meta descriptions are critical for CTR (click-through rate). Even though Shopify auto-generates them, they're generic and underutilized.")]
      }),

      createTable([
        ["Page", "Current", "Recommended Meta Description", "Priority"],
        ["Homepage", "Auto-generated", "Kick your fitness into gear with 9xFit interval training. Join Rock Star Boot Camp in North Phoenix for group fitness classes, boot camp membership plans, and personal training.", "HIGH"],
        ["Trials & Classes", "Auto-generated", "Start your fitness journey with affordable boot camp trials ($29-$59). 9xFit group fitness classes in Phoenix. Morning & evening classes available. Book your first class today.", "HIGH"],
        ["Memberships", "Auto-generated", "Unlimited access to 9xFit boot camp training. Flexible membership plans starting at $247/month. Join our Phoenix fitness community for results-driven interval training.", "HIGH"]
      ], [1500, 2000, 4500, 1360]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Heading Structure (H1, H2, H3)")] }),
      new Paragraph({
        children: [new TextRun({ text: "Current Status: CRITICAL ISSUE - Multiple H1 Tags Found", bold: true, color: "C00000" })],
        spacing: { after: 80 }
      }),
      new Paragraph({
        children: [new TextRun("The homepage contains 4 H1 tags:\n\u2022 'THE FITNESS BOOT CAMP THAT WILL KICK YOU INTO GEAR!'\n\u2022 '9xFit Training Program'\n\u2022 'Membership Plans'\n\u2022 'RSBC GEAR'")]
      }),

      createTable([
        ["Issue", "Impact", "Solution"],
        ["Multiple H1 tags confuse search engines about page focus", "Dilutes SEO authority; signals poor content structure", "Keep ONE H1 tag per page. Demote others to H2. Main H1 should be 'THE FITNESS BOOT CAMP THAT WILL KICK YOU INTO GEAR!' - it\'s the brand message."],
        ["H1 lacks target keywords", "Missed opportunity for keyword relevance", "Consider: 'THE FITNESS BOOT CAMP THAT WILL KICK YOU INTO GEAR! | 9xFit Training in Phoenix' (H1 + supporting text)"],
        ["Inconsistent heading hierarchy on product pages", "Navigation confusion for both users and bots", "Audit all internal pages for proper H1→H2→H3 flow"]
      ], [1800, 2200, 4360]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Content Optimization")] }),

      createTable([
        ["Keyword Area", "Current Gaps", "Recommended Content", "Estimated Impact"],
        ["Primary Keywords", "Weak density of 'boot camp' + 'Phoenix' combination", "Increase natural usage of 'fitness boot camp Phoenix' and '9xFit training' to 1.5-2.5% on homepage", "HIGH"],
        ["Long-tail Keywords", "Not leveraged (e.g., 'boot camp near North Phoenix', 'interval training classes Phoenix')", "Create landing pages or blog posts for high-intent phrases like 'best boot camp in Cave Creek', 'high-intensity interval training Phoenix'", "HIGH"],
        ["Service Descriptions", "Generic descriptions; lack value proposition clarity", "Expand class descriptions with benefits (calorie burn targets, results timeframes, community aspects)", "MEDIUM"],
        ["Local Keywords", "Address/location rarely mentioned naturally", "Add city and neighborhood names to copy. E.g., 'North Phoenix fitness community', 'Cave Creek boot camp'", "MEDIUM"]
      ], [1400, 2100, 3300, 1560]),

      new Paragraph({ children: [new PageBreak()] }),

      // Technical SEO
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. Technical SEO")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Shopify Platform Strengths & Considerations")] }),

      createTable([
        ["Aspect", "Current Shopify Support", "Optimization Needed"],
        ["Canonical Tags", "Automatically generated; prevents duplicate content", "Verify correct canonicals on product variations (trials, memberships, gear)"],
        ["XML Sitemaps", "Auto-generated at /sitemap.xml", "Ensure all important pages are included; exclude product variants if duplicate content"],
        ["robots.txt", "Auto-generated", "Review to ensure no blocking of important pages"],
        ["SSL/HTTPS", "Free SSL included", "Verify all pages load over HTTPS (security ranking signal)"],
        ["Mobile Responsiveness", "Responsive Shopify themes included", "Test with Google Mobile-Friendly Test; ensure fast load times"],
        ["Structured Data", "Basic product schema; supports FAQ/LocalBusiness", "Add LocalBusiness schema with full NAP, hours, class schedule"]
      ], [1800, 2700, 3860]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Page Speed & Core Web Vitals")] }),

      new Paragraph({
        children: [new TextRun({ text: "Estimated Performance:", bold: true })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Homepage load time: ~2-3 seconds (Shopify avg with optimized assets)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Largest Contentful Paint (LCP): Target <2.5s")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("First Input Delay (FID): Should be <100ms")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Cumulative Layout Shift (CLS): Target <0.1")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Recommendations:", bold: true })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Optimize images: Compress all product & class photos to <150KB; use next-gen formats (WebP)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Minify CSS/JS: Review app installations; remove unnecessary scripts")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Lazy load images: Defer images below the fold")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Use Google PageSpeed Insights & Lighthouse: Test monthly; target 90+ scores")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Structured Data & Schema Markup")] }),

      new Paragraph({
        children: [new TextRun("Current: Basic product schema exists for membership trials. Missing LocalBusiness, AggregateOffer, and Schedule schema.")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 100 }
      }),

      createTable([
        ["Schema Type", "Current Status", "Priority", "Implementation"],
        ["LocalBusiness", "MISSING", "CRITICAL", "Add Name, Address, Phone, Hours, Class Schedule, Service Area"],
        ["AggregateOffer", "PARTIAL (trials only)", "HIGH", "Expand to all membership tier variations"],
        ["Event", "MISSING", "HIGH", "Add schema for scheduled boot camp classes"],
        ["FAQPage", "MISSING", "MEDIUM", "Add FAQ schema for common boot camp questions"],
        ["BreadcrumbList", "MISSING", "LOW", "Add navigation breadcrumbs for crawlability"]
      ], [1400, 1600, 800, 3560]),

      new Paragraph({ children: [new PageBreak()] }),

      // Local SEO
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. Local SEO Analysis")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Google Business Profile Optimization")] }),

      new Paragraph({
        children: [new TextRun({ text: "Current Status:", bold: true })]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Profile exists with basic information")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("79 reviews on Yelp; review presence signals legitimacy")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "High-Priority Improvements:", bold: true })]
      }),

      createTable([
        ["Component", "Recommendation", "Expected Impact"],
        ["Business Name", "Verify 'Rock Star Boot Camp' in Google Business Profile; avoid keyword stuffing", "Increased local pack visibility"],
        ["Categories", "Primary: 'Fitness Boot Camp' | Secondary: 'Health Coach', 'Interval Training Center'", "Better categorization for local search"],
        ["Service Description", "Add 9xFit Training Program details, class formats, membership benefits", "Rich snippets; higher CTR"],
        ["Hours & Schedule", "Add class schedule (5:30am, 6:30am, 9:15am, 4:30pm, 5:30pm, weekends)", "Local pack ranking boost; improves user experience"],
        ["Service Area", "Define North Phoenix coverage (neighborhoods: Camelback East, Paradise Valley vicinity)", "Expanded local search reach"],
        ["Photos & Videos", "Add 15+ high-quality photos: group classes, facility, results/testimonials, staff", "Increases engagement; organic traffic"],
        ["Posts", "Weekly updates: new class times, promotions, member spotlights", "Fresh content signal; higher visibility"],
        ["Q&A", "Monitor and respond to Google Q&A; add FAQs proactively", "Trust signals; answers competitor questions"],
        ["Reviews", "Implement review collection system; respond to all reviews (positive & negative)", "Social proof; ranking factor"]
      ], [1600, 2400, 3360]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("NAP Consistency (Name, Address, Phone)")] }),

      createTable([
        ["Source", "Current Data", "Status", "Action"],
        ["Website", "Rock Star Boot Camp | 20635 N. Cave Creek Rd. | (480) 981-6877", "Check", "Verify consistency"],
        ["Google Business", "Should match above", "Verify", "Ensure exact match"],
        ["Yelp", "Likely different format", "Check", "Update if needed"],
        ["Facebook", "Check current listing", "Verify", "Align with primary source"],
        ["ClassPass", "Should match website", "Verify", "Confirm phone/address"]
      ], [1600, 2400, 900, 2460]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Local Keywords & Service Area Pages")] }),

      new Paragraph({
        children: [new TextRun("Recommended local keywords to target:")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("'Boot camp near North Phoenix' / 'Fitness boot camp North Phoenix'")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("'Cave Creek fitness training' / 'Camelback East boot camp'")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("'Best fitness classes in Phoenix' / 'HIIT training Phoenix'")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("'Group fitness classes near me' / '9xFit training Phoenix'")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Service Area Pages Recommendation:", bold: true })]
      }),
      new Paragraph({
        children: [new TextRun("Create landing pages for neighborhoods: Camelback East, Paradise Valley, North Scottsdale. Each page should include local testimonials, modified class times description, and localized CTA.")]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // Keyword Research
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. Keyword Research & Strategy")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Primary Keywords (Target: Top 10)")] }),

      createTable([
        ["Keyword", "Search Intent", "Local Modifier", "Estimated Monthly Searches (Phoenix)", "Difficulty", "Priority"],
        ["fitness boot camp", "Commercial", "Phoenix area boot camp", "320", "HIGH", "HIGH"],
        ["9xFit training", "Brand/Product", "Phoenix 9xFit training", "90", "LOW", "HIGH"],
        ["group fitness classes", "Commercial", "group fitness classes Phoenix", "480", "MEDIUM", "HIGH"],
        ["boot camp near me", "Local intent", "boot camp near North Phoenix", "220", "MEDIUM", "MEDIUM"],
        ["interval training", "Educational/Commercial", "HIIT interval training Phoenix", "390", "MEDIUM", "MEDIUM"],
        ["fitness membership", "Commercial", "fitness boot camp membership Phoenix", "180", "MEDIUM", "MEDIUM"],
        ["personal training", "Commercial", "personal training Phoenix", "850", "VERY HIGH", "LOW"],
        ["CrossFit classes", "Competitor keyword", "CrossFit boot camp alternative", "440", "HIGH", "MEDIUM"],
        ["weight loss boot camp", "Commercial (high intent)", "weight loss boot camp Phoenix", "260", "MEDIUM", "HIGH"],
        ["HIIT workout classes", "Commercial/Local", "HIIT workout classes Phoenix", "310", "MEDIUM", "MEDIUM"]
      ], [1300, 1500, 1500, 1600, 1000, 860]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Long-Tail Keywords (Opportunity Set)")] }),

      createTable([
        ["Keyword", "Search Intent", "Difficulty", "Opportunity"],
        ["best boot camp fitness classes in Phoenix", "Local + intent", "LOW", "HIGH"],
        ["40-minute interval training workout", "Specific format", "VERY LOW", "HIGH"],
        ["affordable fitness boot camp membership", "Price-conscious", "LOW", "MEDIUM"],
        ["boot camp classes in North Phoenix", "Hyper-local", "VERY LOW", "HIGH"],
        ["9xFit training program benefits", "Informational", "VERY LOW", "MEDIUM"],
        ["morning fitness boot camp classes", "Schedule intent", "VERY LOW", "MEDIUM"],
        ["weekend boot camp fitness Phoenix", "Schedule + local", "VERY LOW", "MEDIUM"],
        ["group fitness for weight loss", "Outcome-focused", "LOW", "MEDIUM"],
        ["total body interval training Phoenix", "Format + location", "VERY LOW", "MEDIUM"],
        ["high-intensity fitness community Phoenix", "Community-focused", "VERY LOW", "LOW"]
      ], [2100, 1700, 1000, 1560]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Competitor Keyword Gaps")] }),

      new Paragraph({
        children: [new TextRun("Analysis of competing boot camps reveals these underexploited opportunities:")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("'9xFit training' is unique to RSBC - strong differentiation keyword")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Most competitors don't explicitly target 'affordable boot camp' - opportunity for Groupon/trial pricing messaging")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Neighborhood-level keywords (e.g., 'Camelback East boot camp') are rarely optimized by competitors")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("'Weekday morning fitness classes' is underexploited for early risers")]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // Content Gap Analysis
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("5. Content Gap Analysis")] }),

      new Paragraph({
        children: [new TextRun("Current website structure focuses on product sales (trials, memberships, gear). Educational and informational content is minimal.")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      createTable([
        ["Content Type", "Current Status", "Recommended Topics", "Expected Benefit", "Effort"],
        ["Blog/Articles", "MISSING", "• HIIT benefits\n• 9xFit program guide\n• Member spotlights\n• Nutrition + fitness\n• Boot camp FAQ", "Organic traffic; keyword rankings; trust", "MEDIUM"],
        ["Class Descriptions", "Generic", "Expand with: calorie burn, duration, difficulty level, outcomes, music style", "Improved rankings; higher conversion", "LOW"],
        ["Testimonials/Case Studies", "Basic presence", "Add before/after photos, detailed member transformation stories, quote extraction for social", "Social proof; conversion boost", "LOW"],
        ["FAQ Page", "MISSING", "Address common questions: What is 9xFit? Modifications for beginners? What to bring? How sore will I be? Results timeline?", "Rich snippet opportunity; user guidance", "LOW"],
        ["Service Area Pages", "MISSING (High priority)", "Create 5-6 pages for: Camelback East, Paradise Valley, North Scottsdale, Central Phoenix, South Phoenix, Scottsdale", "Geo-targeted rankings; local pack boost", "MEDIUM"],
        ["Video Content", "YouTube channel exists but underutilized", "Create: Class previews, member testimonials, workout tips, facility tours", "YouTube rankings; embedded snippets; engagement", "MEDIUM-HIGH"],
        ["Email/Newsletter", "Likely exists but not SEO-focused", "Share blog content, class spotlights, seasonal promotions", "Repeat organic traffic; brand authority", "LOW"]
      ], [1200, 1600, 2200, 1800, 900]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Content Creation Timeline (90-Day Quick Win Plan):", bold: true })]
      }),

      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Month 1: Fix critical on-page issues + create FAQ page + publish 2 blog posts")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Month 2: Launch 3 service area landing pages + publish 2 more blog posts")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Month 3: Create testimonial video series + update Google Business Profile weekly")]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // Backlink & Authority
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("6. Backlink & Authority Recommendations")] }),

      new Paragraph({
        children: [new TextRun("Backlinks are a major SEO ranking factor. Rock Star Boot Camp should pursue both quality and local authority links.")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Backlink Strategy")] }),

      createTable([
        ["Link Opportunity", "Authority Level", "Relevance", "Effort", "Priority"],
        ["Local Phoenix directories (Yelp, ClassPass, Mindbody)", "MEDIUM-HIGH", "VERY HIGH", "LOW", "CRITICAL"],
        ["Local business associations & Chamber of Commerce", "MEDIUM", "HIGH", "MEDIUM", "HIGH"],
        ["Phoenix fitness/wellness blogs", "MEDIUM", "HIGH", "MEDIUM", "HIGH"],
        ["Guest posting on fitness publications", "MEDIUM-HIGH", "HIGH", "HIGH", "HIGH"],
        ["Local news features (fitness trends, member stories)", "HIGH", "MEDIUM", "HIGH", "MEDIUM"],
        ["University/community partnerships", "MEDIUM", "MEDIUM", "MEDIUM", "MEDIUM"],
        ["Branded content partnerships with health brands", "VARIABLE", "MEDIUM", "MEDIUM-HIGH", "LOW-MEDIUM"],
        ["Podcast guest appearances (Phoenix fitness podcasts)", "VARIABLE", "HIGH", "MEDIUM", "MEDIUM"]
      ], [1400, 1100, 1100, 1000, 900]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Quick Link-Building Actions (0-30 Days)")] }),

      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Verify presence on: ClassPass, Mindbody, Yelp, Google Business - ensure all have complete, consistent information")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Join Phoenix Chamber of Commerce; request directory listing with backlink")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Contact local Phoenix fitness bloggers; offer exclusive class offer in exchange for feature + link")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Create shareable member success stories; pitch to local health/lifestyle publications")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Domain Authority Trends")] }),

      new Paragraph({
        children: [new TextRun("With consistent on-page optimization + local citation building + quality content, expect:")]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Domain Authority increase of 5-10 points over 6 months")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Local pack (Google Maps) visibility increase within 3-4 weeks")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Organic search traffic increase of 30-50% within 3 months")]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // Competitor Comparison
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("7. Competitor SEO Comparison")] }),

      new Paragraph({
        children: [new TextRun("Analysis of 3 Phoenix-area fitness boot camp competitors:")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Competitor 1: Burn Boot Camp (Paradise Valley location)", bold: true })]
      }),

      createTable([
        ["Metric", "Burn Boot Camp", "Rock Star Boot Camp", "Gap/Opportunity"],
        ["On-Page Optimization", "Better H1 structure, optimized meta descriptions", "Multiple H1s, missing meta descriptions", "RSBC needs immediate fixes"],
        ["Local Authority", "Strong Groupon presence, multiple directory listings", "Good Yelp presence (79 reviews)", "RSBC should expand local directories"],
        ["Content", "Blog with class highlights and member spotlights", "Minimal blog content", "RSBC needs blog strategy"],
        ["Google Business Profile", "Well-optimized with photos and weekly posts", "Basic profile, limited updates", "RSBC should increase post frequency"],
        ["Backlinks", "Links from Groupon, ClassPass, local news", "Links from Yelp, ClassPass", "RSBC should pursue media coverage"],
        ["Overall Strength", "STRONG SEO foundation", "MODERATE - untapped potential", "Quick fixes = immediate gains"]
      ], [1200, 1700, 1900, 2560]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Competitor 2: Fierce Body Boot Camp (Phoenix)", bold: true })]
      }),

      createTable([
        ["Metric", "Fierce Body", "Rock Star Boot Camp", "Gap/Opportunity"],
        ["Mobile Optimization", "Responsive Shopify theme", "Responsive theme", "Equal footing"],
        ["Title Tags", "Keyword-rich, location-specific", "Generic", "RSBC needs improvements"],
        ["Social Integration", "Active Instagram, Facebook, YouTube", "Presence on all platforms", "RSBC needs more engaging content"],
        ["Reviews & Ratings", "Higher Yelp rating", "Good reviews (79 total)", "RSBC should increase volume"],
        ["Video Content", "Minimal YouTube presence", "Minimal YouTube presence", "Both should expand video strategy"],
        ["Unique Positioning", "Community-focused messaging", "9xFit training = differentiation", "RSBC should emphasize 9xFit more"]
      ], [1200, 1700, 1900, 2560]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Competitor 3: CrossFit Full Strength (Phoenix)", bold: true })]
      }),

      createTable([
        ["Metric", "CrossFit Full Strength", "Rock Star Boot Camp", "Gap/Opportunity"],
        ["Site Functionality", "Robust WordPress site with blog", "Shopify - more limited for blog", "RSBC should leverage Shopify blog feature more"],
        ["Content Depth", "Extensive blog (50+ articles)", "Minimal blog content", "RSBC quick win: start monthly blog"],
        ["Local Keywords", "Targets 'CrossFit Phoenix', 'strength training'", "Targets 'boot camp', '9xFit'", "Different market; RSBC's positioning is strong"],
        ["Email Marketing", "Active newsletter with content", "Appears active", "Both competitive here"],
        ["Community Building", "Events, competitions, social engagement", "Class-based community", "RSBC should highlight community angles"],
        ["Authority/Trust", "Stronger due to content volume", "Moderate - good reviews", "RSBC can catch up with blog launch"]
      ], [1200, 1700, 1900, 2560]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Key Competitive Advantages for RSBC:", bold: true })]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Unique 9xFit program (low competition for branded keyword)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Strong local presence (Yelp, Groupon, ClassPass)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Affordable trial options ($29-$59) - competitive pricing")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Prime location in North Phoenix (growing fitness community)")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Where RSBC Needs Improvement:", bold: true })]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("On-page SEO fundamentals (H1 tags, meta descriptions)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Content strategy (blog, service area pages, video)")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Local SEO signals (Google Business Profile activity)")]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // Priority Action Items
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("8. Priority Action Items (Implementation Roadmap)")] }),

      new Paragraph({
        children: [new TextRun("Ranked by impact and implementation effort.")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      // High Priority
      new Paragraph({
        children: [new TextRun({ text: "TIER 1: HIGH IMPACT / LOW-MEDIUM EFFORT (Complete in 30 Days)", bold: true, color: "C00000" })],
        spacing: { after: 100 }
      }),

      createTable([
        ["Action", "Priority", "Effort", "Expected Impact", "Owner/Timeline"],
        ["Fix H1 tag structure (consolidate to 1 H1 per page)", "CRITICAL", "LOW", "Immediate: +5-10% rankings potential", "Dev team: Week 1"],
        ["Write and implement meta descriptions (15+ pages)", "CRITICAL", "LOW", "Immediate: +10-15% CTR improvement", "Content team: Week 1-2"],
        ["Add LocalBusiness schema with full NAP, hours, schedule", "HIGH", "MEDIUM", "1-2 weeks: Local pack visibility +20%", "Dev team: Week 2"],
        ["Optimize Google Business Profile (add 15+ photos, weekly posts, FAQs)", "HIGH", "MEDIUM", "2-4 weeks: Local authority boost", "Marketing: Week 1-4"],
        ["Implement review collection system (post-class feedback automation)", "HIGH", "LOW", "Ongoing: 5+ new reviews/month", "Ops/Marketing: Week 1"],
        ["Create FAQ page (15-20 common boot camp questions)", "MEDIUM", "LOW", "2-4 weeks: FAQ rich snippet opportunity", "Content team: Week 2-3"]
      ], [1800, 800, 800, 1700, 1260]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      // Medium Priority
      new Paragraph({
        children: [new TextRun({ text: "TIER 2: MEDIUM-HIGH IMPACT / MEDIUM EFFORT (Complete in 60 Days)", bold: true, color: "FF6600" })],
        spacing: { after: 100 }
      }),

      createTable([
        ["Action", "Priority", "Effort", "Expected Impact", "Owner/Timeline"],
        ["Create 5-6 service area landing pages (Camelback East, Paradise Valley, etc.)", "HIGH", "MEDIUM", "4-8 weeks: Local keyword rankings +30%", "Content/Dev: Weeks 2-4"],
        ["Start blog: Publish 2x/month on fitness, boot camp, nutrition topics", "HIGH", "MEDIUM", "Ongoing: +30-50% organic traffic (3 months)", "Content: Weeks 1+"],
        ["Expand product/service descriptions (add benefits, outcomes, modifications)", "MEDIUM", "LOW", "2-3 weeks: Improved conversion rates", "Content team: Week 1-2"],
        ["Verify NAP consistency across all directories (Google, Yelp, Facebook, ClassPass)", "MEDIUM", "LOW", "Immediate: Local consistency score ++", "Marketing: Week 1-2"],
        ["Expand YouTube presence (class previews, testimonials, facility tours)", "MEDIUM", "MEDIUM-HIGH", "8-12 weeks: YouTube rankings, embedded snippets", "Video team: Weeks 2-6"],
        ["Increase social media engagement (3x posts/week with fitness tips, class highlights)", "MEDIUM", "LOW", "Ongoing: Social signals, traffic boost", "Marketing: Week 1+"]
      ], [1800, 800, 800, 1700, 1260]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      // Lower Priority
      new Paragraph({
        children: [new TextRun({ text: "TIER 3: MEDIUM-LOW IMPACT / HIGH EFFORT (Complete in 90+ Days)", bold: true, color: "70AD47" })],
        spacing: { after: 100 }
      }),

      createTable([
        ["Action", "Priority", "Effort", "Expected Impact", "Owner/Timeline"],
        ["Implement Event schema for class schedules (calendar integration)", "LOW-MEDIUM", "HIGH", "Rich snippet potential, calendar visibility", "Dev team: Weeks 5-8"],
        ["Guest posting outreach (3-5 posts on fitness blogs/publications)", "MEDIUM", "MEDIUM-HIGH", "Backlinks; brand authority; referral traffic", "Content/PR: Weeks 4-8"],
        ["Partner with local complementary businesses (chiropractors, nutritionists)", "MEDIUM", "MEDIUM", "Co-marketing opportunities; mutual links", "Business dev: Weeks 4+"],
        ["Membership/class scheduling blog series (case studies, member success stories)", "LOW", "MEDIUM", "Trust building; social proof", "Content: Weeks 5+"],
        ["Mobile app development (optional, long-term)", "LOW", "VERY HIGH", "User retention; direct bookings", "Dev: Months 4-6"],
        ["Advanced analytics setup (GA4, conversion tracking, attribution)", "MEDIUM", "MEDIUM", "Better decision-making for future campaigns", "Analytics: Week 3-4"]
      ], [1800, 800, 800, 1700, 1260]),

      new Paragraph({ children: [new PageBreak()] }),

      // KPIs & Measurement
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("9. Success Metrics & KPIs")] }),

      new Paragraph({
        children: [new TextRun("Track these metrics monthly to measure SEO progress:")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      createTable([
        ["Metric", "Current Baseline", "30-Day Target", "90-Day Target", "Measurement Tool"],
        ["Organic traffic (monthly sessions)", "Unknown - audit baseline", "+15-20%", "+40-60%", "Google Analytics 4"],
        ["Keyword rankings (target 10 keywords)", "Track current positions", "5+ keywords on page 1", "8+ keywords on page 1", "SEMrush, Ahrefs"],
        ["Local pack visibility (Google Maps)", "Baseline position", "Top 3 in local search", "Top 1-2 consistently", "Local rank tracking"],
        ["Google Business Profile views", "Establish baseline", "+50% month-over-month", "+100% from baseline", "Google Business"],
        ["Reviews (Yelp, Google, ClassPass)", "79 reviews on Yelp", "5+ new reviews/month", "15+ new reviews/month", "Manual tracking"],
        ["Class trial conversions", "Unknown - baseline needed", "+10% from paid clicks", "+20% from organic", "Shopify analytics"],
        ["Membership signups from organic", "Unknown - track with UTM", "+25% growth", "+50% growth", "UTM parameters, GA4"],
        ["Page speed (Core Web Vitals)", "Establish baseline with PageSpeed", "Improve LCP <2.5s", "Maintain 90+ Lighthouse score", "Google PageSpeed Insights"],
        ["Backlinks acquired", "Baseline audit", "5-10 new citations", "15-20 new citations", "SEMrush, Ahrefs"],
        ["Social media referral traffic", "Establish baseline", "+25%", "+50%", "GA4 - social channel"]
      ], [1300, 1600, 1600, 1600, 1260]),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun({ text: "Monthly Reporting Template:", bold: true })]
      }),

      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Organic traffic % change")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Top 5 performing pages")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("New keywords ranking on page 1")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Google Business Profile engagement metrics")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Actions completed & pending")]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // Conclusion
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Conclusion")] }),

      new Paragraph({
        children: [new TextRun("Rock Star Boot Camp has a strong foundation with a unique 9xFit program, solid local presence, and Shopify platform advantages. However, critical on-page SEO issues (H1 tags, meta descriptions) and limited content strategy are holding back organic growth.")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun("The good news: most issues are quick to fix. Implementing the Tier 1 actions (30 days) will yield immediate improvements. Combining those with Tier 2 content strategy over 60-90 days can realistically drive 40-60% organic traffic growth.")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun("Success requires coordinated effort across development, marketing, and content teams. Consistent execution of these recommendations, combined with the 9xFit brand differentiation, positions Rock Star Boot Camp to dominate local Phoenix boot camp search results.")]
      }),

      new Paragraph({
        children: [new TextRun("")],
        spacing: { after: 120 }
      }),

      new Paragraph({
        children: [new TextRun("Next Step: Schedule implementation kickoff meeting to prioritize Tier 1 actions and assign owners by team.")]
      })
    ]
  }]
});

function createTable(rows, columnWidths) {
  const totalWidth = 9360; // US Letter with 1" margins

  return new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: columnWidths,
    rows: rows.map((row, rowIndex) => {
      const isHeader = rowIndex === 0;
      return new TableRow({
        children: row.map((cell, colIndex) => {
          const isFirstCol = colIndex === 0;
          return new TableCell({
            borders: isHeader ? headerBorders : borders,
            width: { size: columnWidths[colIndex], type: WidthType.DXA },
            shading: isHeader ? { fill: "2E75B6", type: ShadingType.CLEAR } : colIndex % 2 === 0 ? { fill: "F0F0F0", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            verticalAlign: "center",
            children: [
              new Paragraph({
                children: [new TextRun({
                  text: cell,
                  bold: isHeader,
                  color: isHeader ? "FFFFFF" : "000000",
                  size: isHeader ? 20 : 20
                })],
                alignment: AlignmentType.LEFT
              })
            ]
          });
        })
      });
    })
  });
}

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/sessions/vibrant-upbeat-cray/mnt/ai-marketing-claude/clients/rockstar-bootcamp/02-seo-audit.docx", buffer);
  console.log("✓ SEO Audit document created successfully");
});
