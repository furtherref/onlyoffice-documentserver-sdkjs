'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    createSdkHarness,
    createTextElement
} = require('./helpers/load-sdk-source');

// ---------------------------------------------------------------------
// Task 4: resolve spell-check language per word character
// ---------------------------------------------------------------------
//
// These tests load the real production resolver
// (word/Editor/Paragraph/Run/LanguageResolver.js, plus its
// FontClassification/bidi-types/commonDefines dependencies), the real
// word/Editor/SpellChecker/ParagraphCollector.js and the real
// word/Editor/Run.js into a single node:vm sandbox, then exercise the
// actual CParagraphSpellCheckerCollector and ParaRun.prototype.CheckSpelling
// implementations. Nothing about language resolution is mocked.

function createCollectorFixture(options)
{
    options = options || {};

    const harness = createSdkHarness({looseGlobals: true});
    const sandbox = harness.sandbox;

    harness.loadResolverStack();

    // Minimal double for AscWord.CParagraphContentPos: the collector only
    // needs Update(pos, depth) / Get(depth) bookkeeping, nothing about its
    // internal representation is exercised by these tests.
    function CParagraphContentPos()
    {
        this.pos = [];
    }
    CParagraphContentPos.prototype.Update = function(nPos, nDepth)
    {
        this.pos[nDepth] = nPos;
    };
    CParagraphContentPos.prototype.Get = function(nDepth)
    {
        return this.pos[nDepth];
    };
    sandbox.AscWord.CParagraphContentPos = CParagraphContentPos;

    harness.load('word/Editor/SpellChecker/ParagraphCollector.js');
    harness.load('word/Editor/Run.js');

    const additions = [];
    const oSpellChecker = {
        Paragraph: {
            isRtlDirection: function()
            {
                return !!options.isRtl;
            }
        },
        Add: function()
        {
            additions.push(Array.prototype.slice.call(arguments));
        }
    };

    const collector = new sandbox.AscWord.CParagraphSpellCheckerCollector(
        oSpellChecker,
        !!options.isForceFullCheck
    );

    const CheckSpellingFn = sandbox.AscWord.ParaRun.prototype.CheckSpelling;

    function createRun(runOptions)
    {
        runOptions = runOptions || {};
        const content = runOptions.content || [];
        const lang = runOptions.lang || {};

        return {
            Content: content,
            SpellingMarks: [],
            GetReviewType: function()
            {
                return runOptions.reviewType !== undefined
                    ? runOptions.reviewType
                    : sandbox.reviewtype_Common;
            },
            Get_CompiledPr: function()
            {
                return {Lang: lang, Caps: false};
            },
            GetCompiledEastAsiaBeforeDirect: function()
            {
                return runOptions.eastAsiaBeforeDirect;
            },
            IsEmpty: function()
            {
                return content.length === 0;
            },
            CheckSpelling: CheckSpellingFn
        };
    }

    function createElement(elementOptions)
    {
        return createTextElement(elementOptions);
    }

    return {
        collector: collector,
        additions: additions,
        createRun: createRun,
        createElement: createElement,
        sandbox: sandbox
    };
}

function addedWordsAndLcids(additions)
{
    return additions.map(function(item)
    {
        return {word: item[4], lcid: item[5]};
    });
}

test('collector splits one mixed run at a script language boundary',
function()
{
    const fixture = createCollectorFixture();
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    const run = fixture.createRun({
        lang: {Val: 1033, EastAsia: 2052, Bidi: 1025},
        eastAsiaBeforeDirect: 2052,
        content: [
            fixture.createElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            }),
            fixture.createElement({
                codePoint: 0x61,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            })
        ]
    });

    run.CheckSpelling(fixture.collector, 0);
    fixture.collector.FlushWord();

    assert.deepEqual(addedWordsAndLcids(fixture.additions), [
        {word: '中', lcid: 2052},
        {word: 'a', lcid: 1033}
    ]);

    // Non-overlapping positions: the Chinese segment covers [0, 1), the
    // Latin segment covers [1, 2).
    assert.equal(fixture.additions[0][1], 0);
    assert.equal(fixture.additions[0][3], 1);
    assert.equal(fixture.additions[1][1], 1);
    assert.equal(fixture.additions[1][3], 2);
});

test('collector flushes before collecting the first character of a new LCID',
function()
{
    const fixture = createCollectorFixture();
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    const run = fixture.createRun({
        lang: {Val: 1033, EastAsia: 2052, Bidi: 1025},
        eastAsiaBeforeDirect: 2052,
        content: [
            fixture.createElement({
                codePoint: 0x61,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            }),
            fixture.createElement({
                codePoint: 0x62,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            }),
            fixture.createElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });

    run.CheckSpelling(fixture.collector, 0);
    fixture.collector.FlushWord();

    assert.deepEqual(addedWordsAndLcids(fixture.additions), [
        {word: 'ab', lcid: 1033},
        {word: '中', lcid: 2052}
    ]);
    // The Han character starts a fresh word at position 2, the Han
    // position in the run's content.
    assert.equal(fixture.additions[1][1], 2);
});

test('combining mark retains the current word LCID', function()
{
    const fixture = createCollectorFixture();
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;
    const Other = sdk.AscBidi.DIRECTION_FLAG.Other;

    const run = fixture.createRun({
        lang: {Val: 1033, EastAsia: 2052, Bidi: 1025},
        eastAsiaBeforeDirect: 2052,
        content: [
            fixture.createElement({
                codePoint: 0x61,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            }),
            // A neutral combining mark: direction is neither LTR nor RTL,
            // so ResolveRunElementLanguage yields undefined and must not
            // manufacture or flush a language.
            fixture.createElement({
                codePoint: 0x0301,
                direction: Other,
                fontSlot: sdk.AscWord.fontslot_Unknown
            })
        ]
    });

    run.CheckSpelling(fixture.collector, 0);

    // No intermediate flush happened while collecting the two characters.
    assert.equal(fixture.additions.length, 0);
    assert.equal(fixture.collector.CurLcid, 1033);

    fixture.collector.FlushWord();

    assert.deepEqual(addedWordsAndLcids(fixture.additions), [
        {word: 'á', lcid: 1033}
    ]);
});

test('timer resume refreshes language and fallback for the resumed run',
function()
{
    const fixture = createCollectorFixture();
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    // Simulate leftover state from a previous run whose East Asia
    // fallback was Chinese (2052).
    fixture.collector.HandleLang({Val: 1033, EastAsia: 2052, Bidi: 1025},
        2052);

    // The resumed run's own EastAsia value (3084, fr-CA) is not an
    // ideograph language, so resolution must fall back to *this run's*
    // GetCompiledEastAsiaBeforeDirect() (1042, Korean) rather than the
    // stale 2052 left over from the previous run.
    const run = fixture.createRun({
        lang: {Val: 1033, EastAsia: 3084, Bidi: 1025},
        eastAsiaBeforeDirect: 1042,
        content: [
            fixture.createElement({
                codePoint: 0xac00,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });

    fixture.collector.UpdatePos(0, 0);
    fixture.collector.SetFindStart(true);

    run.CheckSpelling(fixture.collector, 0);
    fixture.collector.FlushWord();

    assert.deepEqual(addedWordsAndLcids(fixture.additions), [
        {word: '가', lcid: 1042}
    ]);
});

test('removed run remains excluded from spelling', function()
{
    const fixture = createCollectorFixture();
    const sdk = fixture.sandbox;

    const run = fixture.createRun({
        reviewType: sdk.reviewtype_Remove,
        lang: {Val: 1033, EastAsia: 2052, Bidi: 1025},
        content: [
            fixture.createElement({
                codePoint: 0x61,
                direction: sdk.AscBidi.DIRECTION_FLAG.LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            })
        ]
    });
    // Prove Get_CompiledPr is never even consulted for a removed run.
    run.Get_CompiledPr = function()
    {
        throw new Error('Get_CompiledPr must not be called');
    };

    run.CheckSpelling(fixture.collector, 0);

    assert.equal(fixture.additions.length, 0);
    assert.equal(fixture.collector.CheckedCounter, 0);
    assert.equal(fixture.collector.bWord, false);
});

test('English dictionary request is preserved', function()
{
    const fixture = createCollectorFixture();
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;
    const word = 'wrngword';

    const run = fixture.createRun({
        lang: {Val: 1033, EastAsia: 1033, Bidi: 1025},
        eastAsiaBeforeDirect: 1033,
        content: word.split('').map(function(ch)
        {
            return fixture.createElement({
                codePoint: ch.codePointAt(0),
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            });
        })
    });

    run.CheckSpelling(fixture.collector, 0);
    fixture.collector.FlushWord();

    assert.deepEqual(addedWordsAndLcids(fixture.additions), [
        {word: word, lcid: 1033}
    ]);
});
