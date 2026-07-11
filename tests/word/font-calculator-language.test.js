'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    createSdkHarness,
    createTextElement,
    createTextPr
} = require('./helpers/load-sdk-source');

// ---------------------------------------------------------------------
// Task 5: resolve status-bar language for selections and collapsed
// carets.
// ---------------------------------------------------------------------
//
// These tests load the real resolver stack, the real
// word/Editor/Paragraph/ParagraphContentPos.js, the real
// word/Editor/Paragraph.js (Paragraph.prototype navigation,
// CParagraphRunElements and the new
// Paragraph.prototype.GetNearestStrongTextLanguageContext helper) and
// the real word/Editor/Paragraph/Run/FontCalculator.js
// (CFontCalculator.prototype.Calculate) into a single node:vm sandbox.
//
// AscWord.ParaRun, paragraph run content and paragraph-navigation
// "run" containers are explicit test doubles (FakeParaRun below);
// AscWord.CParagraphContentPos is the real production class.
// docContent.CheckSelectedRunContent is an explicit double that
// invokes the real CFontCalculator callback with declared half-open
// ranges, exactly as the real CDocumentContent implementation would.

function createFontCalculatorFixture(options)
{
    options = options || {};

    const harness = createSdkHarness({looseGlobals: true});
    const sandbox = harness.sandbox;

    // Minimal doubles required only so the two large production files
    // evaluate their top-level class wiring without throwing; neither
    // stub is exercised by any assertion below.
    sandbox.AscCommon.CColor = function() {};
    sandbox.CDocumentContentElementBase = function() {};

    harness.loadResolverStack();
    harness.load('word/Editor/Paragraph/ParagraphContentPos.js');
    harness.load('word/Editor/Paragraph.js');
    harness.load('word/Editor/Paragraph/Run/FontCalculator.js');

    const AscWord = sandbox.AscWord;
    const AscBidi = sandbox.AscBidi;
    const Paragraph = sandbox.Paragraph;
    const CParagraphContentPos = AscWord.CParagraphContentPos;

    assert.equal(typeof Paragraph.prototype.GetNearestStrongTextLanguageContext,
        'function');
    assert.equal(typeof sandbox.AscWord.FontCalculator.Calculate, 'function');

    // ------------------------------------------------------------
    // FakeParaRun: explicit AscWord.ParaRun double. Implements the
    // small production-shaped navigation contract
    // (GetNextRunElements/GetPrevRunElements/Get_ClassesByPos/
    // Get_ParaContentPos) the same way word/Editor/Run.js's real
    // ParaRun does, so the real Paragraph.prototype navigation and
    // CParagraphRunElements bookkeeping this task's helper depends on
    // run unmodified against it. Font/language content accessors
    // (GetElement, GetElementsCount, GetFontSlotInRange,
    // GetCompiledEastAsiaBeforeDirect, Get_CompiledPr, GetParagraph)
    // are plain call-counted doubles.
    // ------------------------------------------------------------
    function FakeParaRun(runOptions)
    {
        runOptions = runOptions || {};
        this._elements = runOptions.elements || [];
        this._textPr = runOptions.textPr || createTextPr({});
        this._paragraph = runOptions.paragraph || null;
        this._eastAsiaBeforeDirect = runOptions.eastAsiaBeforeDirect;
        this._fontSlotByPosition = runOptions.fontSlotByPosition !== undefined
            ? runOptions.fontSlotByPosition
            : AscWord.fontslot_ASCII;
        this._fontSlotInRange = runOptions.fontSlotInRange !== undefined
            ? runOptions.fontSlotInRange
            : AscWord.fontslot_None;
        this._directionFlagInRange = runOptions.directionFlagInRange !== undefined
            ? runOptions.directionFlagInRange
            : AscBidi.DIRECTION_FLAG.LTR;
        this._caretPos = runOptions.caretPos !== undefined
            ? runOptions.caretPos
            : 0;
        this._isParaEndRun = !!runOptions.isParaEndRun;
        this._reviewType = runOptions.reviewType;
        this.calls = {
            GetElement: 0,
            GetElementPositions: [],
            GetElementsCount: 0,
            GetFontSlotInRange: 0,
            GetDirectionFlagInRange: 0,
            GetCompiledEastAsiaBeforeDirect: 0,
            Get_CompiledPr: 0,
            GetParagraph: 0,
            GetNextRunElements: 0,
            GetPrevRunElements: 0
        };
    }
    FakeParaRun.prototype.IsEmpty = function()
    {
        return 0 === this._elements.length;
    };
    FakeParaRun.prototype.IsParaEndRun = function()
    {
        return this._isParaEndRun;
    };
    FakeParaRun.prototype.GetReviewType = function()
    {
        return this._reviewType;
    };
    FakeParaRun.prototype.GetParagraph = function()
    {
        this.calls.GetParagraph++;
        return this._paragraph;
    };
    FakeParaRun.prototype.Get_CompiledPr = function()
    {
        this.calls.Get_CompiledPr++;
        return this._textPr;
    };
    FakeParaRun.prototype.GetParaEndCompiledPr = function()
    {
        return this._textPr;
    };
    FakeParaRun.prototype.GetElementsCount = function()
    {
        this.calls.GetElementsCount++;
        return this._elements.length;
    };
    FakeParaRun.prototype.GetElement = function(nPos)
    {
        this.calls.GetElement++;
        this.calls.GetElementPositions.push(nPos);
        return this._elements[nPos];
    };
    FakeParaRun.prototype.GetFontSlotInRange = function()
    {
        this.calls.GetFontSlotInRange++;
        return this._fontSlotInRange;
    };
    FakeParaRun.prototype.GetDirectionFlagInRange = function()
    {
        this.calls.GetDirectionFlagInRange++;
        return this._directionFlagInRange;
    };
    FakeParaRun.prototype.GetFontSlotByPosition = function()
    {
        return this._fontSlotByPosition;
    };
    FakeParaRun.prototype.GetCompiledEastAsiaBeforeDirect = function()
    {
        this.calls.GetCompiledEastAsiaBeforeDirect++;
        return this._eastAsiaBeforeDirect;
    };
    FakeParaRun.prototype.GetParent = function()
    {
        return null;
    };
    FakeParaRun.prototype.Get_ClassesByPos = function(Classes)
    {
        Classes.push(this);
    };
    FakeParaRun.prototype.Get_ParaContentPos = function(bSelection, bStart, ContentPos)
    {
        ContentPos.Add(this._caretPos);
    };
    // Mirrors word/Editor/Run.js ParaRun.prototype.GetNextRunElements
    // exactly (production contract for CParagraphRunElements), against
    // this double's own `_elements` array instead of a real run's
    // Content.
    FakeParaRun.prototype.GetNextRunElements = function(oRunElements, isUseContentPos, nDepth)
    {
        this.calls.GetNextRunElements++;
        if (true === isUseContentPos)
            oRunElements.SetStartClass(this.GetParent());

        let nStartPos = true === isUseContentPos
            ? oRunElements.ContentPos.Get(nDepth)
            : 0;

        for (let nCurPos = nStartPos, nCount = this._elements.length; nCurPos < nCount; ++nCurPos)
        {
            if (oRunElements.IsEnoughElements() || this.IsEmpty())
                return;

            oRunElements.UpdatePos(nCurPos, nDepth);
            oRunElements.Add(this._elements[nCurPos], this);
        }
    };
    // Mirrors word/Editor/Run.js ParaRun.prototype.GetPrevRunElements.
    FakeParaRun.prototype.GetPrevRunElements = function(oRunElements, isUseContentPos, nDepth)
    {
        this.calls.GetPrevRunElements++;
        if (true === isUseContentPos)
            oRunElements.SetStartClass(this.GetParent());

        let nStartPos = true === isUseContentPos
            ? oRunElements.ContentPos.Get(nDepth) - 1
            : this._elements.length - 1;

        for (let nCurPos = nStartPos; nCurPos >= 0; --nCurPos)
        {
            if (oRunElements.IsEnoughElements() || this.IsEmpty())
                return;

            oRunElements.UpdatePos(nCurPos, nDepth);
            oRunElements.Add(this._elements[nCurPos], this);
        }
    };

    sandbox.AscWord.ParaRun = FakeParaRun;

    function buildRun(runOptions)
    {
        return new FakeParaRun(runOptions);
    }

    function buildParagraph(runs, caretRunIndex)
    {
        let paragraph = Object.create(Paragraph.prototype);
        paragraph.Content = runs;
        paragraph.CurPos  = {ContentPos: caretRunIndex || 0};
        paragraph.Selection = {Use: false};
        paragraph.GetTheme = function()
        {
            return {themeElements: {fontScheme: {}}};
        };
        paragraph.GetNumberingTextPr = function()
        {
            return options.numberingTextPr;
        };
        paragraph.GetLogicDocument = function()
        {
            return {
                IsDocumentEditor: function()
                {
                    return false !== options.isDocumentEditor;
                }
            };
        };
        runs.forEach(function(run)
        {
            run._paragraph = paragraph;
        });
        return paragraph;
    }

    function makePos(runIndex, inRunPos)
    {
        let pos = new CParagraphContentPos();
        pos.Add(runIndex);
        pos.Add(inRunPos);
        return pos;
    }

    function createOutputTextPr()
    {
        return {Lang: {}};
    }

    function createDocContent(runRanges)
    {
        return {
            IsNumberingSelection: function()
            {
                return !!options.isNumberingSelection;
            },
            IsTextSelectionUse: function()
            {
                return !!options.hasSelection;
            },
            IsSelectionEmpty: function()
            {
                return !options.hasSelection;
            },
            GetCurrentParagraph: function()
            {
                return options.currentParagraph;
            },
            CheckSelectedRunContent: function(callback)
            {
                for (let nIndex = 0; nIndex < runRanges.length; ++nIndex)
                {
                    let spec = runRanges[nIndex];
                    let stop = callback(spec.run, spec.start, spec.end);
                    if (stop)
                        break;
                }
            }
        };
    }

    let outputTextPr = createOutputTextPr();
    let runs = (options.runs || []).map(buildRun);
    let paragraph = options.paragraph
        || buildParagraph(runs, options.caretRunIndex);

    if (undefined === options.currentParagraph)
        options.currentParagraph = paragraph;

    let docContent = createDocContent(options.selectionRanges || []);

    return {
        calculator: sandbox.AscWord.FontCalculator,
        docContent: docContent,
        outputTextPr: outputTextPr,
        paragraph: paragraph,
        runs: runs,
        sandbox: sandbox,
        AscWord: AscWord,
        AscBidi: AscBidi,
        buildRun: buildRun,
        buildParagraph: buildParagraph,
        makePos: makePos,
        createElement: createTextElement,
        createTextPr: createTextPr
    };
}

// =======================================================================
// Selection tests
// =======================================================================

test('selection resolves homogeneous Han as 2052', function()
{
    const fixture = createFontCalculatorFixture({hasSelection: true});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    const run = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        fontSlotInRange: sdk.AscWord.fontslot_EastAsia,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });
    run._paragraph = fixture.paragraph;
    fixture.docContent.CheckSelectedRunContent = function(callback)
    {
        callback(run, 0, 1);
    };

    assert.equal(typeof fixture.calculator.Calculate, 'function');
    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, 2052);
});

test('selection resolves homogeneous Latin as 1033', function()
{
    const fixture = createFontCalculatorFixture({hasSelection: true});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    const run = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 1033,
        fontSlotInRange: sdk.AscWord.fontslot_ASCII,
        elements: [
            createTextElement({
                codePoint: 0x61,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            })
        ]
    });
    run._paragraph = fixture.paragraph;
    fixture.docContent.CheckSelectedRunContent = function(callback)
    {
        callback(run, 0, 1);
    };

    assert.equal(typeof fixture.calculator.Calculate, 'function');
    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, 1033);
});

test('selection resolves homogeneous RTL from Bidi', function()
{
    const fixture = createFontCalculatorFixture({hasSelection: true});
    const sdk = fixture.sandbox;
    const RTL = sdk.AscBidi.DIRECTION_FLAG.RTL;

    const run = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 1033,
        fontSlotInRange: sdk.AscWord.fontslot_CS,
        elements: [
            createTextElement({
                codePoint: 0x0627,
                direction: RTL,
                fontSlot: sdk.AscWord.fontslot_CS
            })
        ]
    });
    run._paragraph = fixture.paragraph;
    fixture.docContent.CheckSelectedRunContent = function(callback)
    {
        callback(run, 0, 1);
    };

    assert.equal(typeof fixture.calculator.Calculate, 'function');
    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, 1025);
});

test('selection resolves one mixed Han Latin run as undefined', function()
{
    const fixture = createFontCalculatorFixture({hasSelection: true});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    const run = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        fontSlotInRange: sdk.AscWord.fontslot_EastAsia | sdk.AscWord.fontslot_ASCII,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            }),
            createTextElement({
                codePoint: 0x61,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            })
        ]
    });
    run._paragraph = fixture.paragraph;
    fixture.docContent.CheckSelectedRunContent = function(callback)
    {
        callback(run, 0, 2);
    };

    assert.equal(typeof fixture.calculator.Calculate, 'function');
    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, undefined);
    assert.deepEqual(run.calls.GetElementPositions, [0, 1]);
});

test('selection ignores neutral-only content', function()
{
    const fixture = createFontCalculatorFixture({hasSelection: true});
    const sdk = fixture.sandbox;
    const Other = sdk.AscBidi.DIRECTION_FLAG.Other;

    const run = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 1033,
        fontSlotInRange: sdk.AscWord.fontslot_Unknown,
        elements: [
            createTextElement({
                codePoint: 0x20,
                direction: Other,
                fontSlot: sdk.AscWord.fontslot_Unknown
            }),
            createTextElement({
                codePoint: 0x09,
                direction: Other,
                fontSlot: sdk.AscWord.fontslot_Unknown
            })
        ]
    });
    run._paragraph = fixture.paragraph;
    fixture.docContent.CheckSelectedRunContent = function(callback)
    {
        callback(run, 0, 2);
    };

    assert.equal(typeof fixture.calculator.Calculate, 'function');
    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, undefined);
});

test('selection range is half-open and font mask is language-independent',
function()
{
    const fixture = createFontCalculatorFixture({hasSelection: true});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;
    const Other = sdk.AscBidi.DIRECTION_FLAG.Other;

    const run = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        fontSlotInRange: sdk.AscWord.fontslot_EastAsia,
        elements: [
            createTextElement({
                codePoint: 0x20,
                direction: Other,
                fontSlot: sdk.AscWord.fontslot_Unknown
            }),
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            }),
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            }),
            createTextElement({
                codePoint: 0x09,
                direction: Other,
                fontSlot: sdk.AscWord.fontslot_Unknown
            })
        ]
    });
    run._paragraph = fixture.paragraph;
    fixture.docContent.CheckSelectedRunContent = function(callback)
    {
        callback(run, 1, 3);
    };

    assert.equal(typeof fixture.calculator.Calculate, 'function');
    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, 2052);
    assert.deepEqual(run.calls.GetElementPositions, [1, 2]);
    assert.equal(run.calls.GetFontSlotInRange, 1);
});

// =======================================================================
// Caret tests
// =======================================================================

test('caret language changes without changing current-run font properties',
function()
{
    const fixture = createFontCalculatorFixture({hasSelection: false});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    let currentTextPr = createTextPr({Val: 1033, EastAsia: 1033, Bidi: 1033});
    currentTextPr.Bold = true;
    currentTextPr.Italic = false;
    currentTextPr.FontSize = 14;
    currentTextPr.RFonts.Ascii.Name = 'Arial';

    const runCurrent = fixture.buildRun({
        textPr: currentTextPr,
        fontSlotByPosition: sdk.AscWord.fontslot_ASCII,
        elements: []
    });
    const runNext = fixture.buildRun({
        textPr: createTextPr({Val: 1041, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });

    const paragraph = fixture.buildParagraph([runCurrent, runNext], 0);
    fixture.docContent.GetCurrentParagraph = function()
    {
        return paragraph;
    };

    assert.equal(
        typeof paragraph.GetNearestStrongTextLanguageContext,
        'function'
    );
    assert.equal(typeof fixture.calculator.Calculate, 'function');
    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, 2052);
    assert.equal(fixture.outputTextPr.Bold, true);
    assert.equal(fixture.outputTextPr.Italic, false);
    assert.equal(fixture.outputTextPr.FontSize, 14);
    assert.equal(fixture.outputTextPr.FontFamily.Name, 'Arial');
});

test('caret context keeps element run textPr and fallback atomic', function()
{
    const fixture = createFontCalculatorFixture({hasSelection: false});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    const runPrev = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 1033, Bidi: 1033}),
        eastAsiaBeforeDirect: 1033,
        elements: [
            createTextElement({
                codePoint: 0x62,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            })
        ]
    });
    // A neutral gap run makes runPrev distance 2 so runNext (distance
    // 1, immediately adjacent) wins unambiguously - this test is about
    // not blending textPr across runs, not about the tie rule (see the
    // dedicated tie test below).
    const runGap = fixture.buildRun({
        elements: [
            createTextElement({
                codePoint: 0x20,
                direction: sdk.AscBidi.DIRECTION_FLAG.Other,
                fontSlot: sdk.AscWord.fontslot_Unknown
            })
        ]
    });
    const runCurrent = fixture.buildRun({elements: []});
    const runNext = fixture.buildRun({
        textPr: createTextPr({Val: 9999, EastAsia: 2052, Bidi: 1111}),
        eastAsiaBeforeDirect: 2052,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });

    const paragraph = fixture.buildParagraph(
        [runPrev, runGap, runCurrent, runNext],
        2
    );
    const pos = fixture.makePos(2, 0);

    assert.equal(
        typeof paragraph.GetNearestStrongTextLanguageContext,
        'function'
    );
    const context = paragraph.GetNearestStrongTextLanguageContext(pos);

    assert.ok(context);
    assert.equal(context.Run, runNext);
    assert.equal(context.ResolvedLcid, 2052);
    assert.equal(context.TextPr, runNext._textPr);
});

test('caret chooses closest strong context and previous on a tie', function()
{
    const sdkFixture = createFontCalculatorFixture({hasSelection: false});
    const sdk = sdkFixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    function strongLatinRun(fixture)
    {
        return fixture.buildRun({
            textPr: createTextPr({Val: 1033, EastAsia: 1033, Bidi: 1033}),
            eastAsiaBeforeDirect: 1033,
            elements: [
                createTextElement({
                    codePoint: 0x62,
                    direction: LTR,
                    fontSlot: sdk.AscWord.fontslot_ASCII
                })
            ]
        });
    }
    function strongHanRun(fixture)
    {
        return fixture.buildRun({
            textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
            eastAsiaBeforeDirect: 2052,
            elements: [
                createTextElement({
                    codePoint: 0x4e2d,
                    direction: LTR,
                    fontSlot: sdk.AscWord.fontslot_EastAsia
                })
            ]
        });
    }
    function neutralRun(fixture)
    {
        return fixture.buildRun({
            elements: [
                createTextElement({
                    codePoint: 0x20,
                    direction: sdk.AscBidi.DIRECTION_FLAG.Other,
                    fontSlot: sdk.AscWord.fontslot_Unknown
                })
            ]
        });
    }

    // Closer wins: previous is distance 2 (a neutral run sits between
    // the caret and the Latin run), next is distance 1 (adjacent Han
    // run). Next must win.
    {
        const fixture = createFontCalculatorFixture({hasSelection: false});
        const runPrev    = strongLatinRun(fixture);
        const runGap     = neutralRun(fixture);
        const runCurrent = fixture.buildRun({elements: []});
        const runNext    = strongHanRun(fixture);
        const paragraph = fixture.buildParagraph(
            [runPrev, runGap, runCurrent, runNext],
            2
        );
        const pos = fixture.makePos(2, 0);

        const context = paragraph.GetNearestStrongTextLanguageContext(pos);
        assert.ok(context);
        assert.equal(context.ResolvedLcid, 2052);
        assert.equal(context.Distance, 1);
    }

    // Equal distance: both sides are adjacent (distance 1). Previous
    // must win the tie.
    {
        const fixture = createFontCalculatorFixture({hasSelection: false});
        const runPrev    = strongLatinRun(fixture);
        const runCurrent = fixture.buildRun({elements: []});
        const runNext    = strongHanRun(fixture);
        const paragraph = fixture.buildParagraph(
            [runPrev, runCurrent, runNext],
            1
        );
        const pos = fixture.makePos(1, 0);

        const context = paragraph.GetNearestStrongTextLanguageContext(pos);
        assert.ok(context);
        assert.equal(context.ResolvedLcid, 1033);
        assert.equal(context.Run, runPrev);
    }
});

test('caret handles paragraph start end and empty paragraph', function()
{
    const sdkFixture = createFontCalculatorFixture({hasSelection: false});
    const sdk = sdkFixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    // Paragraph start: only following text exists.
    {
        const fixture = createFontCalculatorFixture({hasSelection: false});
        const runCurrent = fixture.buildRun({elements: []});
        const runNext = fixture.buildRun({
            textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
            eastAsiaBeforeDirect: 2052,
            elements: [
                createTextElement({
                    codePoint: 0x4e2d,
                    direction: LTR,
                    fontSlot: sdk.AscWord.fontslot_EastAsia
                })
            ]
        });
        const paragraph = fixture.buildParagraph([runCurrent, runNext], 0);
        const pos = fixture.makePos(0, 0);

        const context = paragraph.GetNearestStrongTextLanguageContext(pos);
        assert.ok(context);
        assert.equal(context.ResolvedLcid, 2052);
    }

    // Paragraph end: only preceding text exists.
    {
        const fixture = createFontCalculatorFixture({hasSelection: false});
        const runPrev = fixture.buildRun({
            textPr: createTextPr({Val: 1033, EastAsia: 1033, Bidi: 1033}),
            eastAsiaBeforeDirect: 1033,
            elements: [
                createTextElement({
                    codePoint: 0x62,
                    direction: LTR,
                    fontSlot: sdk.AscWord.fontslot_ASCII
                })
            ]
        });
        const runCurrent = fixture.buildRun({elements: []});
        const paragraph = fixture.buildParagraph([runPrev, runCurrent], 1);
        const pos = fixture.makePos(1, 0);

        const context = paragraph.GetNearestStrongTextLanguageContext(pos);
        assert.ok(context);
        assert.equal(context.ResolvedLcid, 1033);
    }

    // Empty paragraph: nothing to find in either direction.
    {
        const fixture = createFontCalculatorFixture({hasSelection: false});
        const runCurrent = fixture.buildRun({elements: []});
        const paragraph = fixture.buildParagraph([runCurrent], 0);
        const pos = fixture.makePos(0, 0);

        const context = paragraph.GetNearestStrongTextLanguageContext(pos);
        assert.equal(context, null);
    }
});

test('caret skips neutral candidates and forward search terminates',
function()
{
    const fixture = createFontCalculatorFixture({hasSelection: false});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;
    const Other = sdk.AscBidi.DIRECTION_FLAG.Other;

    const runCurrent = fixture.buildRun({elements: []});
    const runNeutral1 = fixture.buildRun({
        elements: [
            createTextElement({
                codePoint: 0x20,
                direction: Other,
                fontSlot: sdk.AscWord.fontslot_Unknown
            })
        ]
    });
    const runNeutral2 = fixture.buildRun({
        elements: [
            createTextElement({
                codePoint: 0x09,
                direction: Other,
                fontSlot: sdk.AscWord.fontslot_Unknown
            })
        ]
    });
    const runStrong = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });

    const paragraph = fixture.buildParagraph(
        [runCurrent, runNeutral1, runNeutral2, runStrong],
        0
    );
    const pos = fixture.makePos(0, 0);

    assert.equal(
        typeof paragraph.GetNearestStrongTextLanguageContext,
        'function'
    );
    const context = paragraph.GetNearestStrongTextLanguageContext(pos);

    assert.ok(context);
    assert.equal(context.ResolvedLcid, 2052);

    const totalNavigationCalls = paragraph.Content.reduce(function(sum, run)
    {
        return sum
            + run.calls.GetNextRunElements
            + run.calls.GetPrevRunElements;
    }, 0);
    assert.ok(totalNavigationCalls < 10,
        'navigation must terminate well under the 10-call safety cap, got '
        + totalNavigationCalls);
});

test('review projection excludes a hidden deletion candidate', function()
{
    const fixture = createFontCalculatorFixture({hasSelection: false});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    let hiddenCandidateObserved = false;
    const hiddenElement = createTextElement({
        codePoint: 0x62,
        direction: LTR,
        fontSlot: sdk.AscWord.fontslot_ASCII
    });
    const originalGetDirectionFlag = hiddenElement.GetDirectionFlag;
    hiddenElement.GetDirectionFlag = function()
    {
        hiddenCandidateObserved = true;
        return originalGetDirectionFlag.call(hiddenElement);
    };

    // Final/Original view: BeginViewModeInReview already projected the
    // deletion out, so this fixture's navigation stream never contains
    // hiddenElement at all - it is simply absent from paragraph.Content.
    const runCurrent = fixture.buildRun({elements: []});
    const runVisible = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });

    const paragraph = fixture.buildParagraph([runCurrent, runVisible], 0);
    const pos = fixture.makePos(0, 0);

    assert.equal(
        typeof paragraph.GetNearestStrongTextLanguageContext,
        'function'
    );
    const context = paragraph.GetNearestStrongTextLanguageContext(pos);

    assert.ok(context);
    assert.equal(context.ResolvedLcid, 2052);
    assert.notEqual(context.Element, hiddenElement);
    assert.equal(hiddenCandidateObserved, false);
});

test('Edit mode may use visible deletion markup', function()
{
    const fixture = createFontCalculatorFixture({hasSelection: false});
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;

    const runCurrent = fixture.buildRun({elements: []});
    // Edit/Simple view: deletion markup remains a visible run-element
    // candidate in the navigation stream (reviewtype_Remove is not
    // blanket-rejected by the helper).
    const runDeleted = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        reviewType: sdk.reviewtype_Remove,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });

    const paragraph = fixture.buildParagraph([runCurrent, runDeleted], 0);
    const pos = fixture.makePos(0, 0);

    assert.equal(
        typeof paragraph.GetNearestStrongTextLanguageContext,
        'function'
    );
    const context = paragraph.GetNearestStrongTextLanguageContext(pos);

    assert.ok(context);
    assert.equal(context.ResolvedLcid, 2052);
    assert.equal(context.Run, runDeleted);
});

test('numbering selection keeps numbering language', function()
{
    const fixture = createFontCalculatorFixture({
        isNumberingSelection: true,
        numberingTextPr: {
            Bold: true,
            Italic: false,
            FontSize: 12,
            RFonts: {Ascii: {Name: 'Courier New'}},
            Lang: {Val: 1049}
        }
    });

    assert.equal(typeof fixture.calculator.Calculate, 'function');
    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, 1049);
    assert.equal(fixture.outputTextPr.Bold, true);
    assert.equal(fixture.outputTextPr.FontSize, 12);
});

test('non-Document selection keeps the legacy direction-based language',
function()
{
    const fixture = createFontCalculatorFixture({
        hasSelection: true,
        isDocumentEditor: false
    });
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;
    const run = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        fontSlotInRange:
            sdk.AscWord.fontslot_EastAsia | sdk.AscWord.fontslot_ASCII,
        directionFlagInRange: LTR,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            }),
            createTextElement({
                codePoint: 0x61,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_ASCII
            })
        ]
    });
    run._paragraph = fixture.paragraph;
    fixture.docContent.CheckSelectedRunContent = function(callback)
    {
        callback(run, 0, 2);
    };

    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, 1033);
    assert.equal(run.calls.GetDirectionFlagInRange, 1);
    assert.equal(run.calls.GetElement, 0);
    assert.equal(run.calls.GetCompiledEastAsiaBeforeDirect, 0);
});

test('non-Document caret keeps the legacy current-run language', function()
{
    const fixture = createFontCalculatorFixture({
        hasSelection: false,
        isDocumentEditor: false
    });
    const sdk = fixture.sandbox;
    const LTR = sdk.AscBidi.DIRECTION_FLAG.LTR;
    const run = fixture.buildRun({
        textPr: createTextPr({Val: 1033, EastAsia: 2052, Bidi: 1025}),
        eastAsiaBeforeDirect: 2052,
        fontSlotByPosition: sdk.AscWord.fontslot_EastAsia,
        caretPos: 0,
        elements: [
            createTextElement({
                codePoint: 0x4e2d,
                direction: LTR,
                fontSlot: sdk.AscWord.fontslot_EastAsia
            })
        ]
    });
    const paragraph = fixture.buildParagraph([run], 0);
    fixture.docContent.GetCurrentParagraph = function()
    {
        return paragraph;
    };

    fixture.calculator.Calculate(fixture.docContent, fixture.outputTextPr);

    assert.equal(fixture.outputTextPr.Lang.Val, 1033);
    assert.equal(run.calls.GetCompiledEastAsiaBeforeDirect, 0);
});
