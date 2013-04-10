namespace opening_word_document
{
    public class POSTagger
    {

        // all the pos tages supported
        // CC Coordinating conjunction
        //CD Cardinal number
        //DT Determiner
        //EX Existential there
        //FW Foreign word
        //IN Preposition or subordinating conjunction
        //JJ Adjective
        //JJR Adjective, comparative
        //JJS Adjective, superlative
        //LS List item marker
        //MD Modal
        //NN Noun, singular or mass
        //NNS Noun, plural
        //NNP Proper noun, singular
        //NNPS Proper noun, plural
        //PDT Predeterminer
        //POS Possessive ending
        //PRP Personal pronoun
        //PRP$ Possessive pronoun
        //RB Adverb
        //RBR Adverb, comparative
        //RBS Adverb, superlative
        //RP Particle
        //SYM Symbol
        //TO to
        //UH Interjection
        //VB Verb, base form
        //VBD Verb, past tense
        //VBG Verb, gerund or present participle
        //VBN Verb, past participle
        //VBP Verb, non­3rd person singular present
        //VBZ Verb, 3rd person singular present
        //WDT Wh­determiner
        //WP Wh­pronoun
        //WP$ Possessive wh­pronoun
        //WRB Wh­adverb
        // private static string mModelPath;

        private static OpenNLP.Tools.SentenceDetect.MaximumEntropySentenceDetector mSentenceDetector;
        private static OpenNLP.Tools.Tokenize.EnglishMaximumEntropyTokenizer mTokenizer;
        private static OpenNLP.Tools.PosTagger.EnglishMaximumEntropyPosTagger mPosTagger;
        private static OpenNLP.Tools.Chunker.EnglishTreebankChunker mChunker;
        private static OpenNLP.Tools.Parser.EnglishTreebankParser mParser;
        private static OpenNLP.Tools.NameFind.EnglishNameFinder mNameFinder;
        private static OpenNLP.Tools.Lang.English.TreebankLinker mCoreferenceFinder;


        public static string mModelPath { get; set; }

        public static string[] SplitSentences(string paragraph)
        {
            if (mSentenceDetector == null)
            {
                mSentenceDetector = new OpenNLP.Tools.SentenceDetect.EnglishMaximumEntropySentenceDetector(mModelPath + "EnglishSD.nbin");
            }

            return mSentenceDetector.SentenceDetect(paragraph);
        }

        public static string[] TokenizeSentence(string sentence)
        {
            if (mTokenizer == null)
            {
                mTokenizer = new OpenNLP.Tools.Tokenize.EnglishMaximumEntropyTokenizer(mModelPath + "EnglishTok.nbin");
            }

            return mTokenizer.Tokenize(sentence);
        }

        public static string[] PosTagTokens(string[] tokens)
        {
            if (mPosTagger == null)
            {
                mPosTagger = new OpenNLP.Tools.PosTagger.EnglishMaximumEntropyPosTagger(mModelPath + "EnglishPOS.nbin", mModelPath + @"\Parser\tagdict");
            }

            return mPosTagger.Tag(tokens);
        }

        public static string ChunkSentence(string[] tokens, string[] tags)
        {
            if (mChunker == null)
            {
                mChunker = new OpenNLP.Tools.Chunker.EnglishTreebankChunker(mModelPath + "EnglishChunk.nbin");
            }

            return mChunker.GetChunks(tokens, tags);
        }

        public static OpenNLP.Tools.Parser.Parse ParseSentence(string sentence)
        {
            if (mParser == null)
            {
                mParser = new OpenNLP.Tools.Parser.EnglishTreebankParser(mModelPath, true, false);
            }

            return mParser.DoParse(sentence);
        }

        public static string FindNames(string sentence)
        {
            if (mNameFinder == null)
            {
                mNameFinder = new OpenNLP.Tools.NameFind.EnglishNameFinder(mModelPath + "namefind\\");
            }

            string[] models = new string[] { "date", "location", "money", "organization", "percentage", "person", "time" };
            return mNameFinder.GetNames(models, sentence);
        }

        public static string FindNames(OpenNLP.Tools.Parser.Parse sentenceParse)
        {
            if (mNameFinder == null)
            {
                mNameFinder = new OpenNLP.Tools.NameFind.EnglishNameFinder(mModelPath + "namefind\\");
            }

            string[] models = new string[] { "date", "location", "money", "organization", "percentage", "person", "time" };
            return mNameFinder.GetNames(models, sentenceParse);
        }

        public static string IdentifyCoreferents(string[] sentences)
        {
            if (mCoreferenceFinder == null)
            {
                mCoreferenceFinder = new OpenNLP.Tools.Lang.English.TreebankLinker(mModelPath + "coref");
            }

            System.Collections.Generic.List<OpenNLP.Tools.Parser.Parse> parsedSentences = new System.Collections.Generic.List<OpenNLP.Tools.Parser.Parse>();

            foreach (string sentence in sentences)
            {
                OpenNLP.Tools.Parser.Parse sentenceParse = ParseSentence(sentence);
                string findNames = FindNames(sentenceParse);
                parsedSentences.Add(sentenceParse);
            }
            return mCoreferenceFinder.GetCoreferenceParse(parsedSentences.ToArray());
        }
    }
}
