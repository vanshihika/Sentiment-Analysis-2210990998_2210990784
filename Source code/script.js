const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  PageBreak, LevelFormat, UnderlineType
} = require('docx');
const fs = require('fs');

// ── border helpers ────────────────────────────────────────────────────────────
const bdr   = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
const bdrs  = { top: bdr, bottom: bdr, left: bdr, right: bdr };
const noBdr = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };
const noBdrs= { top: noBdr, bottom: noBdr, left: noBdr, right: noBdr };

const CW = 9360; // content width (US Letter, 1" margins)

// ── cell factories ────────────────────────────────────────────────────────────
function hCell(text, w) {
  return new TableCell({
    borders: bdrs, width: { size: w, type: WidthType.DXA },
    shading: { fill: "1F3864", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20, font: "Times New Roman" })]
    })]
  });
}
function dCell(text, w, shade = "FFFFFF", left = false, bold = false) {
  return new TableCell({
    borders: bdrs, width: { size: w, type: WidthType.DXA },
    shading: { fill: shade, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: left ? AlignmentType.LEFT : AlignmentType.CENTER,
      children: [new TextRun({ text, size: 20, bold, font: "Times New Roman" })]
    })]
  });
}
function noCell(content, w) {
  return new TableCell({
    borders: noBdrs, width: { size: w, type: WidthType.DXA },
    children: content
  });
}

// ── paragraph factories ───────────────────────────────────────────────────────
const sp = (b, a) => ({ before: b, after: a });

function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function centBold(text, size = 24) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: sp(80, 80),
    children: [new TextRun({ text, bold: true, size, font: "Times New Roman" })]
  });
}
function cent(text, size = 22) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: sp(60, 60),
    children: [new TextRun({ text, size, font: "Times New Roman" })]
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: sp(opts.before || 80, opts.after || 80),
    indent: opts.indent ? { left: 720 } : undefined,
    children: [new TextRun({ text, size: 22, bold: opts.bold || false, font: "Times New Roman" })]
  });
}
function secHead(text) {
  return new Paragraph({
    spacing: sp(200, 100),
    children: [new TextRun({ text, bold: true, size: 24, font: "Times New Roman",
      underline: { type: UnderlineType.SINGLE } })]
  });
}
function subHead(text) {
  return new Paragraph({
    spacing: sp(160, 80),
    children: [new TextRun({ text, bold: true, size: 22, font: "Times New Roman" })]
  });
}
function blank(n = 1) {
  return Array.from({ length: n }, () =>
    new Paragraph({ children: [new TextRun("")], spacing: sp(40, 40) })
  );
}
function italicCenter(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: sp(80, 80),
    children: [new TextRun({ text, italics: true, size: 22, font: "Times New Roman" })]
  });
}

// ── numbering ─────────────────────────────────────────────────────────────────
const numbering = {
  config: [{
    reference: "bullets",
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: "\u2022",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  }]
};
function bul(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 }, spacing: sp(40, 40),
    children: [new TextRun({ text, size: 22, font: "Times New Roman" })]
  });
}
function bulB(label, rest) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 }, spacing: sp(40, 40),
    children: [
      new TextRun({ text: label, bold: true,  size: 22, font: "Times New Roman" }),
      new TextRun({ text: rest,               size: 22, font: "Times New Roman" })
    ]
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// DOCUMENT
// ══════════════════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering,
  styles: {
    default: { document: { run: { font: "Times New Roman", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Times New Roman" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [

      // ════════════════════════════════════════════════════════════════
      // COVER PAGE
      // ════════════════════════════════════════════════════════════════
      ...blank(2),
      centBold("PROJECT REPORT OF INDUSTRY ORIENTED HANDS ON EXPERIENCE (IOHE)", 26),
      centBold("ON", 24),
      ...blank(1),
      centBold("A Sentiment-Based Analysis of Online Customer Reviews and Their Influence on Product Ratings", 24),
      ...blank(2),
      cent("submitted in partial fulfilment of the requirements for the award of degree of"),
      ...blank(1),
      centBold("BACHELOR OF ENGINEERING", 24),
      cent("In"),
      centBold("COMPUTER SCIENCE AND ENGINEERING", 24),
      ...blank(2),

      new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: [4680, 4680],
        rows: [new TableRow({ children: [
          noCell([
            new Paragraph({ children: [new TextRun({ text: "Submitted by:", bold: true, size: 22, font: "Times New Roman" })] }),
            new Paragraph({ children: [new TextRun({ text: "Vanshika Choudhary (2210990998)", size: 22, font: "Times New Roman" })] }),
            new Paragraph({ children: [new TextRun({ text: "Sanchita Dutt (2210990784)",      size: 22, font: "Times New Roman" })] }),
          ], 4680),
          noCell([
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Supervised By:", bold: true, size: 22, font: "Times New Roman" })] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Dr. Monika Aggarwal",  size: 22, font: "Times New Roman" })] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Assistant Professor",  size: 22, font: "Times New Roman" })] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Chitkara University",  size: 22, font: "Times New Roman" })] }),
          ], 4680),
        ]})]
      }),

      ...blank(2),
      centBold("DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING", 22),
      ...blank(1),
      centBold("CHITKARA UNIVERSITY INSTITUTE OF ENGINEERING AND TECHNOLOGY", 22),
      centBold("CHITKARA UNIVERSITY, PUNJAB, INDIA", 22),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // CONTENTS
      // ════════════════════════════════════════════════════════════════
      centBold("CONTENTS", 24),
      ...blank(1),
      new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: [1200, 6360, 1800],
        rows: [
          new TableRow({ children: [hCell("S.No.", 1200), hCell("Title", 6360), hCell("Page No.", 1800)] }),
          new TableRow({ children: [dCell("",   1200,"F5F5F5"), dCell("Declaration",       6360,"F5F5F5",true), dCell("3", 1800,"F5F5F5")] }),
          new TableRow({ children: [dCell("",   1200),          dCell("Acknowledgement",   6360,"FFFFFF",true), dCell("4", 1800)] }),
          new TableRow({ children: [dCell("",   1200,"F5F5F5"), dCell("Abstract",          6360,"F5F5F5",true), dCell("5", 1800,"F5F5F5")] }),
          new TableRow({ children: [dCell("1.", 1200),          dCell("Introduction",      6360,"FFFFFF",true), dCell("6\u20137", 1800)] }),
          new TableRow({ children: [dCell("2.", 1200,"F5F5F5"), dCell("Methodology",       6360,"F5F5F5",true), dCell("8\u201310",1800,"F5F5F5")] }),
          new TableRow({ children: [dCell("3.", 1200),          dCell("Tools and Technologies", 6360,"FFFFFF",true), dCell("11\u201312",1800)] }),
          new TableRow({ children: [dCell("4.", 1200,"F5F5F5"), dCell("Implementation",    6360,"F5F5F5",true), dCell("13\u201314",1800,"F5F5F5")] }),
          new TableRow({ children: [dCell("5.", 1200),          dCell("Major Findings / Outcomes / Output / Results", 6360,"FFFFFF",true), dCell("15\u201317",1800)] }),
          new TableRow({ children: [dCell("6.", 1200,"F5F5F5"), dCell("Conclusion and Future Scope", 6360,"F5F5F5",true), dCell("18\u201319",1800,"F5F5F5")] }),
          new TableRow({ children: [dCell("",   1200),          dCell("References",        6360,"FFFFFF",true), dCell("20", 1800)] }),
        ]
      }),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // DECLARATION
      // ════════════════════════════════════════════════════════════════
      centBold("DECLARATION", 24),
      ...blank(1),
      body('I hereby certify that the work which is being presented in the project report entitled \u201cA Sentiment-Based Analysis of Online Customer Reviews and Their Influence on Product Ratings\u201d in partial fulfilment of requirement for the award of the degree of Bachelor of Engineering (Computer Science and Engineering) submitted in the department of Computer Science and Engineering at Chitkara University Institute of Engineering and Technology, Chitkara University, Punjab, India, is an authentic record of my own work carried out under the supervision of Dr. Monika Aggarwal. The matter presented in this project report has not been submitted in any other university/institute for the award of any degree.'),
      ...blank(2),

      new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: [4680, 4680],
        rows: [new TableRow({ children: [
          noCell([
            new Paragraph({ children: [new TextRun({ text: "Place: Rajpura",    size: 22, font: "Times New Roman" })] }),
            new Paragraph({ children: [new TextRun({ text: "Date: 30.04.2026",  size: 22, font: "Times New Roman" })] }),
          ], 4680),
          noCell([
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Vanshika Choudhary, Sanchita Dutt", size: 22, font: "Times New Roman" })] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "(2210990998, 2210990784)",           size: 22, font: "Times New Roman" })] }),
          ], 4680),
        ]})]
      }),

      ...blank(2),
      body("This is to certify that the above statement made by the candidate is correct to the best of my knowledge and belief."),
      ...blank(2),
      new Paragraph({ children: [new TextRun({ text: "Dr. Monika Aggarwal",       bold: true, size: 22, font: "Times New Roman" })] }),
      new Paragraph({ children: [new TextRun({ text: "Assistant Professor",                    size: 22, font: "Times New Roman" })] }),
      new Paragraph({ children: [new TextRun({ text: "Department of Computer Science and Engineering", size: 22, font: "Times New Roman" })] }),
      new Paragraph({ children: [new TextRun({ text: "Chitkara University Institute of Engineering and Technology,", size: 22, font: "Times New Roman" })] }),
      new Paragraph({ children: [new TextRun({ text: "Chitkara University, Punjab, India",      size: 22, font: "Times New Roman" })] }),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // ACKNOWLEDGEMENT
      // ════════════════════════════════════════════════════════════════
      centBold("ACKNOWLEDGEMENT", 24),
      ...blank(1),
      body("We would like to express our sincere gratitude to all those who supported and guided us throughout the completion of this report."),
      body("We are deeply thankful to our faculty mentors and the Department of Computer Science, Chitkara University, for providing us with the academic foundation and resources necessary to undertake this research. Their constructive feedback and encouragement at every stage of this project were invaluable."),
      body("We extend our appreciation to the open-source community, particularly the contributors behind the NLTK library, the VADER Sentiment Analyser (Hutto & Gilbert, 2014), the TextBlob library (Loria, 2018), and the Amazon Fine Food Reviews dataset (McAuley & Leskovec, 2013), whose publicly available tools and data made this empirical study possible."),
      body("We also thank our peers and fellow students for their continuous motivation and insightful discussions. Their perspectives helped sharpen our understanding of Natural Language Processing concepts, correlation analysis, and classifier evaluation methodology."),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // ABSTRACT
      // ════════════════════════════════════════════════════════════════
      centBold("ABSTRACT", 24),
      ...blank(1),
      body("People lean on product reviews before buying online, yet nobody can read millions of them by hand. This report tackles a straightforward question: can automated sentiment analysis actually capture how customers feel, and do the resulting scores line up with the star ratings those same customers assign? We ran 10,000 reviews from the Amazon Fine Food Reviews dataset through VADER (Valence Aware Dictionary and Sentiment Reasoner) and compared the output against the numerical ratings."),
      body("The correlation was moderate and positive (Spearman\u2019s \u03c1 = 0.371, p < 0.001), real enough to be useful but nowhere near a one-to-one relationship. Pearson\u2019s r = 0.504 (r\u00b2 = 25.4%). One pattern stood out: one-star reviews had by far the widest spread in sentiment scores (\u03c3 = 0.637), meaning dissatisfied customers write in very different ways \u2014 sometimes even sounding positive despite handing out the lowest possible rating."),
      body("Three comparison systems were also evaluated \u2014 a majority-class baseline, TextBlob, and a Logistic Regression classifier \u2014 to see where VADER sits on the accuracy-versus-effort scale. VADER achieved perfect accuracy on our silver-standard evaluation set, though this result needs careful interpretation: because the annotation set was constructed partly using VADER agreement, VADER had a built-in advantage on that specific test. Logistic Regression at 84.0% macro F1 is more likely to generalise. Across all systems, negative reviews were the hardest category to classify correctly."),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // 1. INTRODUCTION
      // ════════════════════════════════════════════════════════════════
      secHead("1.  Introduction"),
      body("Before buying something online, most people check the star rating first. It\u2019s fast and simple. But a single number can\u2019t capture everything: a customer might be furious about late shipping but perfectly happy with the food itself \u2014 and still give one star. Another person might genuinely dislike the product and write a scathing review. Same rating, completely different situation. If a machine learning system treats star ratings as ground-truth sentiment labels, it\u2019s going to train on noise it can\u2019t see."),
      body("The gap between what a rating says and what the review text actually says is the core problem this study set out to examine. The two signals are related, but assuming they\u2019re interchangeable causes predictable errors \u2014 and those errors tend to concentrate in the negative class."),
      body("Most existing sentiment analysis work either collapses everything into binary positive/negative without examining all five rating tiers, or assumes star ratings make reliable sentiment labels without questioning that assumption. Very little prior work examines how sentiment variance shifts across the 1\u20135 scale \u2014 which, as this study demonstrates, is where the most interesting pattern is hiding. Evaluation studies in this area also rarely include a majority-class baseline, making accuracy figures difficult to contextualise."),
      body("This report addresses two primary research questions:"),
      bul("Can automated VADER sentiment scores meaningfully capture how customers feel, as reflected in their written reviews?"),
      bul("Do VADER sentiment scores correlate with the numerical star ratings customers assign, and to what extent?"),
      ...blank(1),
      body("Three hypotheses guide the analysis:"),
      bul("H1: VADER compound scores are positively and significantly correlated with star ratings at the individual review level (Spearman\u2019s \u03c1 > 0, p < 0.001)."),
      bul("H2: Sentiment variance differs across rating groups \u2014 low-rating reviews show substantially higher dispersion than high-rating ones."),
      bul("H3: Logistic Regression outperforms VADER on macro F1, and VADER outperforms TextBlob, consistent with their respective levels of sophistication and data requirements."),
      ...blank(1),
      body("To investigate these questions, a subset of 10,000 reviews from the Amazon Fine Food Reviews Dataset [1] is analysed. VADER sentiment scores are compared against star ratings using Spearman and Pearson correlation, and a four-way classifier comparison (majority-class baseline, TextBlob, VADER, Logistic Regression) is conducted on a silver-standard evaluation set."),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // 2. METHODOLOGY
      // ════════════════════════════════════════════════════════════════
      secHead("2.  Methodology"),
      body("The study was designed as a five-stage sequential pipeline: dataset selection, preprocessing, sentiment scoring, classifier evaluation, and statistical correlation analysis."),

      subHead("2.1  Dataset"),
      body("The Amazon Fine Food Reviews dataset [1] was used, covering food product reviews from 2003 to 2012. Each entry includes the review text, a 1\u20135 star rating, helpfulness votes, and product metadata. The dataset is widely used in the NLP field, which facilitates comparison with earlier work. Its age (over a decade old) is acknowledged as a potential limitation for generalisability to contemporary review language."),

      subHead("2.2  Sampling"),
      body("A random sample of 10,000 reviews was drawn using a fixed seed (seed = 42) to ensure reproducibility. The natural class distribution was preserved rather than forcing a balance, as documenting the imbalance was itself one objective of the study."),

      subHead("2.3  Preprocessing"),
      body("Each review was processed through the following sequential steps:"),
      bul("Lowercase conversion of all text."),
      bul("Removal of HTML tags, URLs, and special characters via regular expressions."),
      bul("Word-level tokenisation."),
      bul("Stop-word removal using the NLTK stop-word list (applied only for TF-IDF features; original text preserved for VADER and TextBlob)."),
      bul("Removal of duplicate entries and records with missing text or rating fields."),
      ...blank(1),
      body("Punctuation was deliberately retained for VADER processing, as exclamation marks and capitalisation are part of how VADER scores sentiment intensity [2]."),

      subHead("2.4  Sentiment Scoring"),
      body("VADER was applied via the vaderSentiment library. Each review receives a compound score S \u2208 [\u22121, +1] computed as:"),
      italicCenter("S = x / \u221a(x\u00b2 + \u03b1\u00b2),   \u03b1 = 15"),
      body("Here, x is the sum of token valence scores after VADER\u2019s heuristic rules are applied. TextBlob was also applied, giving a polarity score P \u2208 [\u22121, +1] from a WordNet-based pattern library, with simpler negation handling and no punctuation weighting."),

      subHead("2.5  Sentiment Classification"),
      body("Standard VADER threshold conditions were applied to classify each review:"),
      ...blank(1),
      new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: [2600, 2600, 4160],
        rows: [
          new TableRow({ children: [hCell("Category", 2600), hCell("Threshold", 2600), hCell("Interpretation", 4160)] }),
          new TableRow({ children: [dCell("Positive", 2600,"E8F5E9"), dCell("S \u2265 0.05", 2600,"E8F5E9"), dCell("Review expresses overall positive sentiment", 4160,"E8F5E9",true)] }),
          new TableRow({ children: [dCell("Negative", 2600,"FFF0F0"), dCell("S \u2264 \u22120.05",2600,"FFF0F0"), dCell("Review expresses overall negative sentiment", 4160,"FFF0F0",true)] }),
          new TableRow({ children: [dCell("Neutral",  2600,"FFFDE7"), dCell("|S| < 0.05", 2600,"FFFDE7"), dCell("Ambiguous or sentiment-neutral",             4160,"FFFDE7",true)] }),
        ]
      }),
      ...blank(1),

      subHead("2.6  Building the Annotation Set"),
      body("Since no human-annotated data was available, a 500-review silver-standard evaluation set was constructed using a rating-based proxy labelling approach, following established distant supervision methods [22, 23, 24]:"),
      bulB("Positive (n = 250): ", "rating \u2208 {4, 5} AND VADER score \u2265 0.05"),
      bulB("Negative (n = 250): ", "rating \u2208 {1, 2} AND VADER score \u2264 \u22120.05"),
      bulB("Excluded: ", "three-star reviews and any review where rating and VADER polarity pointed in opposite directions."),
      ...blank(1),
      body("The set was deliberately balanced at 250 reviews per class. A key limitation is that because VADER polarity was used as one of the selection criteria, VADER holds a structural advantage in the subsequent classifier evaluation \u2014 this is addressed explicitly in Section 5."),

      subHead("2.7  Classifier Setup"),
      body("Four systems were compared on the 500-review annotation set:"),
      bulB("Majority-class baseline: ", "predicts \u201cpositive\u201d for every review regardless of content; negative recall is zero by definition."),
      bulB("TextBlob: ", "applied zero-shot using P > 0.05 for positive, P < \u22120.05 for negative."),
      bulB("VADER: ", "applied zero-shot using its standard thresholds on all 500 reviews."),
      bulB("Logistic Regression: ", "trained on an 80/20 stratified split (400 training, 100 test) using TF-IDF features (max 5,000 unigrams and bigrams, minimum document frequency = 2), L2 regularisation (C = 1.0, lbfgs solver). Bootstrap resampling (1,000 iterations) provided 95% confidence intervals. McNemar\u2019s test [27] compared VADER and Logistic Regression on the overlapping test set."),

      subHead("2.8  Correlation Analysis"),
      body("Spearman\u2019s rank correlation (\u03c1) served as the primary association measure, given that star ratings are ordinal \u2014 the step from one star to two does not necessarily equal the step from four to five [16, 17]. Pearson\u2019s r was also computed as a supplementary figure for comparison with earlier work. Both received 95% bootstrap confidence intervals. Kruskal-Wallis H-test compared sentiment score distributions across all five rating groups, and Mann-Whitney U with Bonferroni correction handled pairwise comparisons."),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // 3. TOOLS AND TECHNOLOGIES
      // ════════════════════════════════════════════════════════════════
      secHead("3.  Tools and Technologies"),

      subHead("3.1  Programming Language"),
      body("Python 3.x was used as the primary programming language, chosen for its extensive NLP ecosystem, rapid prototyping capabilities, and widespread adoption in academic research."),

      subHead("3.2  NLP and Sentiment Analysis"),
      bulB("NLTK (Natural Language Toolkit): ", "Used for tokenisation and stop-word removal during the preprocessing stage."),
      bulB("VADER (vaderSentiment): ", "The core sentiment scoring tool [2]. A rule-based, lexicon-driven analyser designed for social-media-style informal text. Requires no labelled training data, runs in real time, and is sensitive to punctuation, capitalisation, and degree modifiers."),
      bulB("TextBlob: ", "A simpler lexicon-based tool [25] included as a zero-shot lower bound to contextualise VADER\u2019s performance."),

      subHead("3.3  Data Handling"),
      bulB("Pandas: ", "Used for loading, cleaning, and managing the dataset in DataFrame format, including null-value removal, deduplication, and column selection."),
      bulB("NumPy: ", "Provided numerical computation support for aggregating scores and computing means, standard deviations, and confidence intervals."),

      subHead("3.4  Machine Learning"),
      bulB("Scikit-learn: ", "Used to implement the Logistic Regression classifier with TF-IDF vectorisation, stratified train/test splitting, L2 regularisation, and bootstrap confidence interval computation."),

      subHead("3.5  Statistical Analysis"),
      bulB("SciPy: ", "The stats module was used to compute Spearman\u2019s \u03c1, Pearson\u2019s r, Kruskal-Wallis H, and Mann-Whitney U statistics with associated p-values. McNemar\u2019s test was applied to compare classifier performance."),

      subHead("3.6  Data Visualisation"),
      bulB("Matplotlib: ", "Used to produce the scatter plot of VADER compound scores versus star ratings, with group-mean overlays and a regression line."),
      bulB("Seaborn: ", "Used for the bar chart showing the frequency distribution of star ratings across the dataset."),

      subHead("3.7  Dataset"),
      bulB("Amazon Fine Food Reviews Dataset [1]: ", "A publicly available collection of food product reviews with metadata, originally compiled by McAuley and Leskovec (2013) and hosted on the Stanford SNAP repository and Kaggle. Its diverse vocabulary, informal writing style, and broad rating distribution make it an appropriate benchmark for lexicon-based NLP evaluation."),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // 4. IMPLEMENTATION
      // ════════════════════════════════════════════════════════════════
      secHead("4.  Implementation"),

      subHead("4.1  Data Loading and Sampling"),
      body("The Amazon Fine Food Reviews Dataset was loaded into a Pandas DataFrame. A random sample of 10,000 records was drawn with a fixed seed (seed = 42) to balance computational efficiency with statistical representativeness. The review text column and the numerical star-rating column (1\u20135) were retained for analysis."),

      subHead("4.2  Preprocessing Pipeline"),
      body("The preprocessing function was applied element-wise to the review text column in the following order:"),
      bul("Convert all text to lowercase (str.lower())."),
      bul("Remove HTML tags and special characters using Python\u2019s re module."),
      bul("Tokenise using NLTK\u2019s word_tokenize()."),
      bul("Remove English stop words using NLTK\u2019s stop-word corpus."),
      bul("Reconstruct the cleaned text by joining filtered tokens."),
      bul("Drop rows with null values (dropna()) and duplicate entries (drop_duplicates())."),
      ...blank(1),
      body("Note: the original (un-tokenised) text was preserved separately for VADER and TextBlob scoring, as punctuation is functionally important to both tools."),

      subHead("4.3  Sentiment Scoring"),
      body("VADER\u2019s SentimentIntensityAnalyzer was instantiated and polarity_scores() was applied to each review. The compound score \u2014 the normalised aggregate polarity index \u2014 was stored as a new column. TextBlob\u2019s .sentiment.polarity attribute was similarly computed and stored for comparison."),

      subHead("4.4  Sentiment Classification"),
      body("A classification function was applied to the compound score column: Positive if S \u2265 0.05; Negative if S \u2264 \u22120.05; Neutral if |S| < 0.05. The resulting label was stored in a dedicated category column."),

      subHead("4.5  Building the Annotation Set"),
      body("The 500-review silver-standard set was constructed by filtering the preprocessed dataset: positive examples required rating \u2208 {4, 5} AND VADER score \u2265 0.05; negative examples required rating \u2208 {1, 2} AND VADER score \u2264 \u22120.05. Reviews with conflicting signals (e.g., high rating but negative VADER score) were excluded. The resulting set was balanced at 250 examples per class and used as the shared evaluation set for all four classifiers."),

      subHead("4.6  Classifier Training and Evaluation"),
      body("The Logistic Regression classifier was trained on the 400-review training split using TF-IDF features (max 5,000 features, unigrams and bigrams, min document frequency = 2) and evaluated on the 100-review held-out test split. Macro F1, precision, and recall were computed for each class. Bootstrap resampling (1,000 iterations) generated 95% confidence intervals for accuracy. McNemar\u2019s test compared VADER and Logistic Regression on the overlapping test examples."),

      subHead("4.7  Correlation Analysis"),
      body("Spearman\u2019s \u03c1 and Pearson\u2019s r were computed between VADER compound scores and star ratings across all 10,000 reviews using scipy.stats. Kruskal-Wallis H and Mann-Whitney U tests (with Bonferroni correction) assessed distributional differences across the five rating groups. Group-level statistics (n, mean, standard deviation, 95% CI) were tabulated for each star-rating tier."),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // 5. RESULTS
      // ════════════════════════════════════════════════════════════════
      secHead("5.  Major Findings / Outcomes / Output / Results"),

      subHead("5.1  Sentiment Distribution Across the Sample"),
      body("After preprocessing, the 10,000 reviews were classified as follows:"),
      ...blank(1),
      new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: [4680, 2340, 2340],
        rows: [
          new TableRow({ children: [hCell("Category", 4680), hCell("n", 2340), hCell("%", 2340)] }),
          new TableRow({ children: [dCell("VADER Sentiment", 4680,"D9E8F5",true,true), dCell("", 2340,"D9E8F5"), dCell("", 2340,"D9E8F5")] }),
          new TableRow({ children: [dCell("Positive (S \u2265 0.05)", 4680,"FFFFFF",true), dCell("8,813", 2340), dCell("88.1%", 2340)] }),
          new TableRow({ children: [dCell("Negative (S \u2264 \u22120.05)", 4680,"FFF5F5",true), dCell("972", 2340,"FFF5F5"), dCell("9.7%", 2340,"FFF5F5")] }),
          new TableRow({ children: [dCell("Neutral", 4680,"FFFFFF",true), dCell("215", 2340), dCell("2.1%", 2340)] }),
          new TableRow({ children: [dCell("Star Rating", 4680,"D9E8F5",true,true), dCell("", 2340,"D9E8F5"), dCell("", 2340,"D9E8F5")] }),
          new TableRow({ children: [dCell("5\u2605", 4680,"FFFFFF",true), dCell("6,383", 2340), dCell("63.8%", 2340)] }),
          new TableRow({ children: [dCell("4\u2605", 4680,"F5F5F5",true), dCell("1,384", 2340,"F5F5F5"), dCell("13.8%", 2340,"F5F5F5")] }),
          new TableRow({ children: [dCell("3\u2605", 4680,"FFFFFF",true), dCell("793",   2340), dCell("7.9%",  2340)] }),
          new TableRow({ children: [dCell("2\u2605", 4680,"F5F5F5",true), dCell("521",   2340,"F5F5F5"), dCell("5.2%",  2340,"F5F5F5")] }),
          new TableRow({ children: [dCell("1\u2605", 4680,"FFFFFF",true), dCell("919",   2340), dCell("9.2%",  2340)] }),
        ]
      }),
      ...blank(1),
      body("A key observation: 88.1% of reviews received a positive VADER score, yet only 63.8% were five-star ratings. That roughly 25-point gap confirms these two signals are not measuring the same thing \u2014 a substantial number of lower-rated reviews still score as positive text."),

      subHead("5.2  H1: Correlation Between Sentiment and Rating"),
      body("Spearman\u2019s \u03c1 = 0.371 (95% CI [0.351, 0.389], p < 0.001). Pearson\u2019s r = 0.504 (95% CI [0.484, 0.524], p < 0.001), r\u00b2 = 25.4%. By Cohen\u2019s conventions this falls in the medium-effect range. H1 is supported."),
      ...blank(1),
      new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: [3200, 1760, 2800, 1600],
        rows: [
          new TableRow({ children: [hCell("Statistic", 3200), hCell("Value", 1760), hCell("95% CI", 2800), hCell("p", 1600)] }),
          new TableRow({ children: [dCell("Spearman\u2019s \u03c1 (primary)",        3200,"FFFFFF",true), dCell("0.371", 1760), dCell("[0.351, 0.389]", 2800), dCell("< 0.001", 1600)] }),
          new TableRow({ children: [dCell("Pearson\u2019s r (supplementary)",     3200,"F5F5F5",true), dCell("0.504", 1760,"F5F5F5"), dCell("[0.484, 0.524]", 2800,"F5F5F5"), dCell("< 0.001", 1600,"F5F5F5")] }),
          new TableRow({ children: [dCell("Explained Variance (r\u00b2)",          3200,"FFFFFF",true), dCell("25.4%", 1760), dCell("\u2014", 2800), dCell("\u2014", 1600)] }),
          new TableRow({ children: [dCell("Effect size (Cohen)",                  3200,"F5F5F5",true), dCell("Medium",1760,"F5F5F5"), dCell("\u2014", 2800,"F5F5F5"), dCell("\u2014", 1600,"F5F5F5")] }),
        ]
      }),
      ...blank(1),
      body("The moderate strength of the association matters. Explaining only about a quarter of rating variance means the two signals are related but not interchangeable. The remaining 75% of variance comes from outside the review text \u2014 delivery, price, brand expectations, and so on."),

      subHead("5.3  H2: Sentiment Variance Across Rating Groups"),
      body("Kruskal-Wallis confirmed that score distributions differ meaningfully across all five groups (H = 1743.3, df = 4, p < 0.001). All pairwise Mann-Whitney comparisons between adjacent rating tiers were significant after Bonferroni correction (all p < 0.001). H2 is supported."),
      ...blank(1),
      new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: [1260, 1260, 1440, 1440, 2520, 1440],
        rows: [
          new TableRow({ children: [hCell("Rating",1260), hCell("n",1260), hCell("Mean (\u03bc)",1440), hCell("SD (\u03c3)",1440), hCell("95% CI",2520), hCell("Cohen\u2019s d vs 5\u2605",1440)] }),
          new TableRow({ children: [dCell("1\u2605",1260,"FFF0F0"), dCell("919",1260,"FFF0F0"), dCell("0.044",1440,"FFF0F0"), dCell("0.637",1440,"FFF0F0"), dCell("[0.002, 0.086]",2520,"FFF0F0"), dCell("1.48 (large)",1440,"FFF0F0")] }),
          new TableRow({ children: [dCell("2\u2605",1260),          dCell("521",1260),          dCell("0.265",1440),          dCell("0.583",1440),          dCell("[0.215, 0.315]",2520),          dCell("0.92 (large)",1440)] }),
          new TableRow({ children: [dCell("3\u2605",1260,"F5F5F5"), dCell("793",1260,"F5F5F5"), dCell("0.512",1440,"F5F5F5"), dCell("0.499",1440,"F5F5F5"), dCell("[0.477, 0.547]",2520,"F5F5F5"), dCell("0.55 (medium)",1440,"F5F5F5")] }),
          new TableRow({ children: [dCell("4\u2605",1260),          dCell("1,384",1260),        dCell("0.720",1440),          dCell("0.364",1440),          dCell("[0.701, 0.739]",2520),          dCell("0.22 (small)",1440)] }),
          new TableRow({ children: [dCell("5\u2605",1260,"E8F5E9"), dCell("6,383",1260,"E8F5E9"), dCell("0.781",1440,"E8F5E9"), dCell("0.304",1440,"E8F5E9"), dCell("[0.773, 0.789]",2520,"E8F5E9"), dCell("\u2014",1440,"E8F5E9")] }),
        ]
      }),
      ...blank(1),
      body("Three findings stand out:"),
      bulB("Finding 1 \u2014 One-star reviews are not, on average, written negatively: ", "Mean VADER score = 0.044. These reviews average out to near-neutral territory despite carrying the lowest possible rating. A large proportion of dissatisfied customers do not express frustration in the angry language VADER picks up on."),
      bulB("Finding 2 \u2014 One-star reviews have the highest variance (\u03c3 = 0.637, Cohen\u2019s d = 1.48 vs five-star): ", "Some one-star writers are clearly furious; others sound measured or even complimentary about parts of the product. Both still give one star."),
      bulB("Finding 3 \u2014 Variance decreases consistently as ratings climb: ", "Five-star reviews are both the most positive and the most linguistically uniform. Positive labels are inherently cleaner than negative ones for training classifiers."),

      subHead("5.4  H3: Four-System Classifier Comparison"),
      ...blank(1),
      new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: [2100, 1260, 960, 840, 840, 840, 840, 1620],
        rows: [
          new TableRow({ children: [hCell("System",2100), hCell("Accuracy",1260), hCell("Macro F1",960), hCell("Pos. Prec.",840), hCell("Pos. Rec.",840), hCell("Neg. Prec.",840), hCell("Neg. Rec.",840), hCell("Training Data",1620)] }),
          new TableRow({ children: [dCell("Majority class",2100,"FFFFFF",true), dCell("50.0%",1260), dCell("0.333",960), dCell("0.50",840), dCell("1.00",840), dCell("\u2014",840), dCell("0.00",840), dCell("None",1620)] }),
          new TableRow({ children: [dCell("TextBlob",2100,"F5F5F5",true), dCell("84.2%",1260,"F5F5F5"), dCell("0.833",960,"F5F5F5"), dCell("0.84",840,"F5F5F5"), dCell("0.83",840,"F5F5F5"), dCell("0.68",840,"F5F5F5"), dCell("0.61",840,"F5F5F5"), dCell("None",1620,"F5F5F5")] }),
          new TableRow({ children: [dCell("VADER *",2100,"FFFDE7",true), dCell("100.0% *",1260,"FFFDE7"), dCell("1.000 *",960,"FFFDE7"), dCell("1.00",840,"FFFDE7"), dCell("1.00",840,"FFFDE7"), dCell("1.00",840,"FFFDE7"), dCell("1.00",840,"FFFDE7"), dCell("None",1620,"FFFDE7")] }),
          new TableRow({ children: [dCell("Logistic Regression",2100,"E8F5E9",true), dCell("84.0%",1260,"E8F5E9"), dCell("0.840",960,"E8F5E9"), dCell("0.92",840,"E8F5E9"), dCell("0.90",840,"E8F5E9"), dCell("0.82",840,"E8F5E9"), dCell("0.79",840,"E8F5E9"), dCell("400 labeled",1620,"E8F5E9")] }),
        ]
      }),
      ...blank(1),
      new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "* VADER\u2019s perfect score reflects the annotation set construction method, not universal superiority (see Section 5.3 and Limitations).", italics: true, size: 20, font: "Times New Roman" })] }),
      new Paragraph({ spacing: { before: 40, after: 60 }, children: [new TextRun({ text: "LR bootstrap 95% CI: [77.0%, 91.0%]. McNemar\u2019s test (VADER vs. LR): \u03c7\u00b2 = 14.06, p < 0.001.", italics: true, size: 20, font: "Times New Roman" })] }),
      ...blank(1),
      body("The majority-class result of 50% on the balanced evaluation set confirms the evaluation design is fair. On the full unbalanced corpus, this same predictor would score 88.1% while catching zero negative reviews \u2014 demonstrating why the balanced design matters."),
      body("The +8.3 percentage-point improvement from TextBlob to Logistic Regression reflects the value of 400 domain-specific labeled examples: the model learns food-specific vocabulary, brand complaints, and texture descriptions that neither word list captures. Negative recall at 0.79 is the best in the table, but still means roughly one in five negative reviews slips through even on the high-confidence cases in the evaluation set."),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // 6. CONCLUSION AND FUTURE SCOPE
      // ════════════════════════════════════════════════════════════════
      secHead("6.  Conclusion and Future Scope"),

      subHead("6.1  Conclusion"),
      body("This study began with a practical question: does VADER actually capture how customers feel, and is that sentiment reflected in the ratings they assign? The answer is yes, but only partially."),
      body("Spearman\u2019s \u03c1 = 0.371 (Pearson\u2019s r = 0.504, p < 0.001) represents a real and meaningful association, but 25.4% explained variance also means the majority of what determines a star rating has nothing to do with the review text. At the individual level, the signals are related but clearly not the same thing."),
      body("The more interesting finding was the variance pattern. One-star reviews spread across almost the entire sentiment scale \u2014 mean of 0.044, standard deviation of 0.637, Cohen\u2019s d = 1.48 versus five-star. Unhappy customers write inconsistently. Five-star reviews, by contrast, are both the most positive and the most linguistically uniform (\u03c3 = 0.304). That asymmetry has a direct practical consequence: positive sentiment labels derived from star ratings are cleaner than negative ones, and that shows up in every classifier\u2019s negative recall."),
      body("On the classifier comparison: VADER\u2019s perfect score on the silver-standard set is a product of how that set was built, not evidence of universal superiority. Logistic Regression at 84.0% macro F1 is the more generalisable result, and its negative recall improvement over TextBlob (0.79 vs. 0.61) is where the 400-label investment actually pays off."),
      body("Three contributions of this study are worth noting:"),
      bul("A Spearman-based correlation analysis using statistics appropriate for ordinal data, giving a more conservative and accurate estimate than Pearson-only prior work."),
      bul("A characterisation of within-tier sentiment variance across all five rating levels, revealing that negative labels are substantially noisier than positive ones."),
      bul("A four-way classifier comparison that includes a majority-class baseline and explicitly accounts for the evaluation bias introduced by silver-standard label construction."),

      subHead("6.2  Future Scope"),
      bulB("Human-Annotated Gold-Standard Evaluation: ", "Annotating a representative subset of reviews \u2014 including the hard cases (ironic, mixed-sentiment, sarcastic) \u2014 would provide a proper gold-standard benchmark and a much more honest picture of real-world classifier performance."),
      bulB("Transformer Model Comparison: ", "A comparison with fine-tuned BERT or RoBERTa [5, 10] would quantify the accuracy\u2013efficiency trade-off more precisely and test whether contextual models better handle the one-star variance problem."),
      bulB("Aspect-Based Sentiment Analysis: ", "Applying aspect-based NLP [11, 12] to decompose reviews into product-specific attributes (taste, packaging, delivery, value) could explain why rating and review text so often point in different directions."),
      bulB("Class Rebalancing: ", "Applying SMOTE [14] or cost-sensitive learning to address the 88.1% positive-class dominance would likely improve negative and neutral recall in production settings."),
      bulB("Cross-Domain and Cross-Platform Extension: ", "Testing the methodology on electronics, healthcare, or hospitality reviews, and on platforms such as Yelp or Google Reviews, would assess how far the findings generalise beyond food products on Amazon."),

      pb(),

      // ════════════════════════════════════════════════════════════════
      // REFERENCES
      // ════════════════════════════════════════════════════════════════
      centBold("REFERENCES", 24),
      ...blank(1),
      ...[
        "[1] J. J. McAuley and J. Leskovec, \u201cFrom amateurs to connoisseurs: Modeling the evolution of user expertise through online reviews,\u201d in Proc. 22nd Int. Conf. World Wide Web (WWW), Rio de Janeiro, Brazil, May 2013, pp. 897\u2013908.",
        "[2] C. J. Hutto and E. Gilbert, \u201cVADER: A parsimonious rule-based model for sentiment analysis of social media text,\u201d in Proc. 8th Int. AAAI Conf. Weblogs and Social Media (ICWSM), Ann Arbor, MI, Jun. 2014.",
        "[3] B. Liu, Sentiment Analysis and Opinion Mining. San Rafael, CA: Morgan & Claypool, 2012.",
        "[4] B. Pang and L. Lee, \u201cOpinion mining and sentiment analysis,\u201d Found. Trends Inf. Retr., vol. 2, no. 1\u20132, pp. 1\u2013135, 2008.",
        "[5] J. Devlin, M.-W. Chang, K. Lee, and K. Toutanova, \u201cBERT: Pre-training of deep bidirectional transformers for language understanding,\u201d in Proc. NAACL-HLT, Minneapolis, MN, Jun. 2019, pp. 4171\u20134186.",
        "[6] S. Baccianella, A. Esuli, and F. Sebastiani, \u201cSentiWordNet 3.0: An enhanced lexical resource for sentiment analysis and opinion mining,\u201d in Proc. LREC, Valletta, Malta, 2010, pp. 2200\u20132204.",
        "[7] N. Hu, P. A. Pavlou, and J. Zhang, \u201cCan online reviews reveal a product\u2019s true quality?\u201d in Proc. 7th ACM Conf. Electronic Commerce (EC), Ann Arbor, MI, 2006, pp. 324\u2013330.",
        "[8] J. A. Chevalier and D. Mayzlin, \u201cThe effect of word of mouth on sales: Online book reviews,\u201d J. Marketing Res., vol. 43, no. 3, pp. 345\u2013354, Aug. 2006.",
        "[9] B. Pang, L. Lee, and S. Vaithyanathan, \u201cThumbs up? Sentiment classification using machine learning techniques,\u201d in Proc. EMNLP, Philadelphia, PA, 2002, pp. 79\u201386.",
        "[10] Y. Liu et al., \u201cRoBERTa: A robustly optimized BERT pretraining approach,\u201d arXiv:1907.11692, 2019.",
        "[11] C. Sun, L. Huang, and X. Qiu, \u201cUtilizing BERT for aspect-based sentiment analysis via constructing auxiliary sentence,\u201d in Proc. NAACL-HLT, Minneapolis, MN, 2019.",
        "[12] W. Zhang, X. Li, Y. Deng, L. Bing, and W. Lam, \u201cA survey on aspect-based sentiment analysis,\u201d IEEE Trans. Knowl. Data Eng., vol. 35, no. 11, pp. 11019\u201311038, 2022.",
        "[13] S. M. Mudambi and D. Schuff, \u201cWhat makes a helpful online review?\u201d MIS Quarterly, vol. 34, no. 1, pp. 185\u2013200, 2010.",
        "[14] N. V. Chawla, K. W. Bowyer, L. O. Hall, and W. P. Kegelmeyer, \u201cSMOTE: Synthetic minority over-sampling technique,\u201d J. Artif. Intell. Res., vol. 16, pp. 321\u2013357, 2002.",
        "[15] M. T. Ribeiro, S. Singh, and C. Guestrin, \u201c\u2018Why should I trust you?\u2019: Explaining the predictions of any classifier,\u201d in Proc. KDD, San Francisco, CA, 2016.",
        "[16] S. Jamieson, \u201cLikert scales: How to (ab)use them,\u201d Med. Educ., vol. 38, no. 12, pp. 1217\u20131218, 2004.",
        "[17] G. Norman, \u201cLikert scales, levels of measurement and the \u2018laws\u2019 of statistics,\u201d Adv. Health Sci. Educ., vol. 15, no. 5, pp. 625\u2013632, 2010.",
        "[18] A. Fradkin, E. Grewal, and D. Holtz, \u201cRecipricity and the platformization of exchange,\u201d Amer. Econ. Rev., vol. 111, no. 3, pp. 765\u2013795, 2021.",
        "[19] V. Sanh, L. Debut, J. Chaumond, and T. Wolf, \u201cDistilBERT, a distilled version of BERT,\u201d arXiv:1910.01108, 2019.",
        "[20] B. Liang et al., \u201cAspect-based sentiment analysis via affective knowledge enhanced graph convolutional networks,\u201d Knowl.-Based Syst., vol. 235, 107643, 2022.",
        "[21] S. Wang and C. D. Manning, \u201cBaselines and bigrams: Simple, good sentiment and topic classification,\u201d in Proc. ACL, 2012, pp. 90\u201394.",
        "[22] A. Go, R. Bhayani, and L. Huang, \u201cTwitter sentiment classification using distant supervision,\u201d Stanford Tech. Rep., 2009.",
        "[23] J. Blitzer, M. Dredze, and F. Pereira, \u201cBiographies, Bollywood, boom-boxes and blenders: Domain adaptation for sentiment classification,\u201d in Proc. ACL, 2007, pp. 440\u2013447.",
        "[24] J. Read, \u201cUsing emoticons to reduce dependency in machine learning techniques for sentiment classification,\u201d in Proc. ACL Student Research Workshop, 2005, pp. 43\u201348.",
        "[25] S. Loria, TextBlob Documentation, 2018. [Online]. Available: https://textblob.readthedocs.io",
        "[26] L. Yue, W. Chen, X. Li, W. Zuo, and M. Yin, \u201cA survey of sentiment analysis in social media,\u201d Knowl. Inf. Syst., vol. 60, no. 2, pp. 617\u2013663, 2019.",
        "[27] T. G. Dietterich, \u201cApproximate statistical tests for comparing supervised classification learning algorithms,\u201d Neural Comput., vol. 10, no. 7, pp. 1895\u20131923, 1998.",
        "[28] T. Mullen and N. Collier, \u201cSentiment analysis using support vector machines with diverse information sources,\u201d in Proc. EMNLP, 2004, pp. 412\u2013418.",
      ].map(ref => new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { before: 40, after: 40 },
        indent: { left: 360, hanging: 360 },
        children: [new TextRun({ text: ref, size: 20, font: "Times New Roman" })]
      }))
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("IOHE_Report_ResearchPaper.docx", buf);
  console.log("Done");
}).catch(e => console.error(e));