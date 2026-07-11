'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
    createSdkHarness,
    createTextElement,
    createTextPr
} = require('./helpers/load-sdk-source');

function loadResolver() {
    const harness = createSdkHarness();
    harness.loadResolverStack();
    assert.equal(
        typeof harness.sandbox.AscWord.ResolveTextLanguage,
        'function'
    );
    assert.equal(
        typeof harness.sandbox.AscWord.ResolveRunElementLanguage,
        'function'
    );
    return harness.sandbox;
}

test('resolver routes EastAsia ASCII and RTL to separate language slots',
function() {
    const sdk = loadResolver();
    const lang = Object.freeze({Val: 1033, EastAsia: 2052, Bidi: 1025});
    const resolve = sdk.AscWord.ResolveTextLanguage;
    const flags = sdk.AscBidi.DIRECTION_FLAG;

    assert.equal(resolve(
        lang,
        sdk.AscWord.fontslot_EastAsia,
        flags.LTR
    ), 2052);
    assert.equal(resolve(
        lang,
        sdk.AscWord.fontslot_ASCII,
        flags.LTR
    ), 1033);
    assert.equal(resolve(
        lang,
        sdk.AscWord.fontslot_HAnsi,
        flags.LTR
    ), 1033);
    assert.equal(resolve(
        lang,
        sdk.AscWord.fontslot_CS,
        flags.RTL
    ), 1025);
    assert.equal(resolve(
        lang,
        sdk.AscWord.fontslot_ASCII,
        flags.Other
    ), undefined);
    assert.equal(resolve(
        lang,
        sdk.AscWord.fontslot_Unknown,
        flags.LTR
    ), undefined);
    assert.equal(resolve(
        lang,
        sdk.AscWord.fontslot_None,
        flags.LTR
    ), undefined);
});

test('resolver uses only a recognized pre-direct EastAsia fallback',
function() {
    const sdk = loadResolver();
    const flags = sdk.AscBidi.DIRECTION_FLAG;
    const resolve = sdk.AscWord.ResolveTextLanguage;

    assert.equal(resolve(
        {Val: 1033, EastAsia: 3084, Bidi: 1025},
        sdk.AscWord.fontslot_EastAsia,
        flags.LTR,
        2052
    ), 2052);
    assert.equal(resolve(
        {Val: 1033, EastAsia: 3084, Bidi: 1025},
        sdk.AscWord.fontslot_EastAsia,
        flags.LTR,
        3084
    ), 3084);
    assert.equal(resolve(
        {Val: 1033, Bidi: 1025},
        sdk.AscWord.fontslot_EastAsia,
        flags.LTR
    ), 1033);
});

test('resolver derives every ideograph LCID from existing SDK exports',
function() {
    const sdk = loadResolver();
    assert.equal(typeof sdk.AscWord.IsIdeographLanguage, 'function');
    const expected = [
        2052, 1066, 1042, 1041, 0x0004, 1028, 3076,
        4100, 5124, 31748, 30724, 42, 17, 18
    ];

    expected.forEach(function(lcid) {
        assert.equal(sdk.AscWord.IsIdeographLanguage(lcid), true);
    });
    assert.equal(sdk.AscWord.IsIdeographLanguage(0x7fff), false);
});

test('resolver does not synthesize entries absent from the SDK map',
function() {
    const sdk = loadResolver();
    assert.equal(typeof sdk.AscWord.IsIdeographLanguage, 'function');

    // An arbitrary regional tag absent from g_oLcidNameToIdMap must not
    // be invented by the base-tag derivation.
    assert.equal(Object.prototype.hasOwnProperty.call(
        sdk.Asc.g_oLcidNameToIdMap,
        'zh-XX'
    ), false);

    // Representative non-member LCIDs that exist in the SDK map but are
    // not ideograph languages must all be rejected, proving the private
    // set contains exactly the 14 derived members and nothing more.
    const nonMembers = [
        1033,  // en-US
        3084,  // fr-CA
        1025,  // ar-SA
        sdk.Asc.g_oLcidNameToIdMap['en'],
        sdk.Asc.g_oLcidNameToIdMap['fr'],
        sdk.Asc.g_oLcidNameToIdMap['ar']
    ];
    nonMembers.forEach(function(lcid) {
        assert.equal(sdk.AscWord.IsIdeographLanguage(lcid), false);
    });
});

test('run element adapter rejects neutral and non-text candidates',
function() {
    const sdk = loadResolver();
    const textPr = createTextPr({
        Val: 1033,
        EastAsia: 2052,
        Bidi: 1025
    });

    assert.equal(sdk.AscWord.ResolveRunElementLanguage(
        createTextElement({
            codePoint: 0x4e2d,
            direction: sdk.AscBidi.DIRECTION_FLAG.LTR,
            fontSlot: sdk.AscWord.fontslot_EastAsia
        }),
        textPr
    ), 2052);
    assert.equal(sdk.AscWord.ResolveRunElementLanguage(
        createTextElement({
            direction: sdk.AscBidi.DIRECTION_FLAG.Other,
            fontSlot: sdk.AscWord.fontslot_Unknown
        }),
        textPr
    ), undefined);
    assert.equal(sdk.AscWord.ResolveRunElementLanguage(
        createTextElement({isText: false}),
        textPr
    ), undefined);
    assert.equal(sdk.AscWord.ResolveRunElementLanguage(
        null,
        textPr
    ), undefined);
    assert.equal(sdk.AscWord.ResolveRunElementLanguage(
        createTextElement({
            direction: sdk.AscBidi.DIRECTION_FLAG.LTR,
            fontSlot: sdk.AscWord.fontslot_ASCII
        }),
        null
    ), undefined);
    assert.equal(sdk.AscWord.ResolveRunElementLanguage(
        createTextElement({
            direction: sdk.AscBidi.DIRECTION_FLAG.LTR,
            fontSlot: sdk.AscWord.fontslot_ASCII
        }),
        {Val: 1033}
    ), undefined);
});

test('resolver never mutates SDK catalogs or caller language objects',
function() {
    const sdk = loadResolver();
    const listBefore = JSON.stringify(
        sdk.Asc.availableIdeographLanguages
    );
    const mapBefore = JSON.stringify(sdk.Asc.g_oLcidNameToIdMap);
    const lang = Object.freeze({Val: 1033, EastAsia: 2052, Bidi: 1025});

    sdk.AscWord.ResolveTextLanguage(
        lang,
        sdk.AscWord.fontslot_EastAsia,
        sdk.AscBidi.DIRECTION_FLAG.LTR
    );

    assert.equal(
        JSON.stringify(sdk.Asc.availableIdeographLanguages),
        listBefore
    );
    assert.equal(JSON.stringify(sdk.Asc.g_oLcidNameToIdMap), mapBefore);
    assert.deepEqual(lang, {Val: 1033, EastAsia: 2052, Bidi: 1025});
});

// ---------------------------------------------------------------------
// Task 3: run-owned pre-direct East Asia snapshot (Run.js)
// ---------------------------------------------------------------------
//
// These tests load the real word/Editor/Run.js source into the VM
// harness (with looseGlobals so unrelated legacy globals resolve to a
// harmless stub) and exercise the actual ParaRun.prototype methods:
// the constructor, Internal_Compile_Pr, Get_CompiledPr and the new
// GetCompiledEastAsiaBeforeDirect accessor. Collaborators the run
// depends on (Paragraph, Styles, MathPrp, the TextPr class itself and
// the compiled-TextPr cache) are provided as small explicit doubles;
// none of ParaRun's own prototype methods used by these tests are
// reimplemented.

function createRunHarnessContext() {
    const harness = createSdkHarness({looseGlobals: true});
    const sandbox = harness.sandbox;

    const mergeLog = [];
    const tags = new WeakMap();
    function tag(obj, name) {
        if (obj) {
            tags.set(obj, name);
        }
        return obj;
    }
    function tagName(obj) {
        return tags.has(obj) ? tags.get(obj) : 'unknown';
    }

    // Explicit CTextPr double: InitDefault seeds a default language,
    // Copy propagates a tag so the initial paragraph-inherited layer
    // shows up in the merge log, and Merge records the layer's tag and
    // only overwrites language slots the layer actually defines.
    function CTextPr() {
        this.Lang = {};
        this.RFonts = {
            Ascii: {Name: 'Arial'},
            CS: {Name: 'Arial'},
            EastAsia: {Name: 'SimSun'},
            HAnsi: {Name: 'Arial'}
        };
        this.FontFamily = {Name: 'Arial', Index: -1};
        this.FontSize = 11;
        this.FontSizeCS = 11;
        this.FontScale = 100;
        this.Bold = false;
        this.BoldCS = false;
        this.Italic = false;
        this.ItalicCS = false;
    }
    CTextPr.prototype.InitDefault = function() {
        this.Lang = {Val: 1033, EastAsia: 1033, Bidi: 1033};
    };
    CTextPr.prototype.Copy = function() {
        if (tags.has(this)) {
            mergeLog.push(tagName(this));
        }
        const copy = new CTextPr();
        copy.Lang = Object.assign({}, this.Lang);
        return copy;
    };
    CTextPr.prototype.Merge = function(layer) {
        if (!layer) {
            return;
        }
        mergeLog.push(tagName(layer));
        if (layer.Lang) {
            if (layer.Lang.Val !== undefined) {
                this.Lang.Val = layer.Lang.Val;
            }
            if (layer.Lang.EastAsia !== undefined) {
                this.Lang.EastAsia = layer.Lang.EastAsia;
            }
            if (layer.Lang.Bidi !== undefined) {
                this.Lang.Bidi = layer.Lang.Bidi;
            }
        }
    };
    CTextPr.prototype.ReplaceThemeFonts = function() {};
    CTextPr.prototype.CheckFontScale = function() {};

    sandbox.CTextPr = CTextPr;
    sandbox.AscWord.CTextPr = CTextPr;
    sandbox.para_Run = 'run';
    sandbox.para_Math_Run = 'math_run';
    sandbox.styletype_Character = 'character';

    function CStylesDouble() {}
    sandbox.AscWord.CStyles = CStylesDouble;

    // Cache double: deduplicates by the final serialized language only
    // (mirrors g_textPrCache's job of merging equal compiled TextPr
    // objects); it never stores or looks at any run-owned scalar.
    const cacheStore = new Map();
    const cache = {
        add: function(textPr) {
            const key = JSON.stringify({Lang: textPr.Lang});
            if (cacheStore.has(key)) {
                return cacheStore.get(key);
            }
            cacheStore.set(key, textPr);
            return textPr;
        },
        remove: function(textPr) {
            for (const [key, value] of cacheStore) {
                if (value === textPr) {
                    cacheStore.delete(key);
                    break;
                }
            }
        }
    };
    sandbox.AscWord.g_textPrCache = cache;

    sandbox.CParagraphContentWithContentBase = function() {};

    // Minimal collaborators touched unconditionally by the constructor.
    sandbox.AscCommon.g_oIdCounter = {Get_NewId: function() { return 1; }};
    sandbox.AscCommon.CContentChanges = function() {};
    sandbox.AscCommon.g_oTableId = {Add: function() {}};

    // Explicit double so RecalcInfo.TextPr is a real mutable boolean
    // flag (the invalidate/recompile test relies on flipping it).
    function CParaRunRecalcInfo() {
        this.TextPr = true;
    }
    sandbox.CParaRunRecalcInfo = CParaRunRecalcInfo;

    harness.load('word/Editor/Run.js');

    const ParaRun = sandbox.AscWord.ParaRun;
    assert.equal(typeof ParaRun, 'function');
    assert.equal(typeof ParaRun.prototype.Internal_Compile_Pr, 'function');
    assert.equal(typeof ParaRun.prototype.Get_CompiledPr, 'function');

    return {
        sandbox: sandbox,
        ParaRun: ParaRun,
        CTextPr: CTextPr,
        cache: cache,
        mergeLog: mergeLog,
        tag: tag
    };
}

function buildRun(ctx, options) {
    const tag = ctx.tag;

    const styleTextPr = tag(new ctx.CTextPr(), 'style');
    if (options.styleEastAsia !== undefined) {
        styleTextPr.Lang = {EastAsia: options.styleEastAsia};
    }
    const stylesInstance = new ctx.sandbox.AscWord.CStyles();
    stylesInstance.Get_Pr = function() {
        return {TextPr: styleTextPr};
    };

    const paraParent = {
        Get_Styles: function() {
            return options.styles === null ? null : stylesInstance;
        }
    };

    const inheritedTextPr = tag(new ctx.CTextPr(), 'paragraph');
    if (options.paragraphEastAsia !== undefined) {
        inheritedTextPr.Lang = {EastAsia: options.paragraphEastAsia};
    }

    const paragraph = options.paragraphMissing ? undefined : {
        Get_CompiledPr2: function() {
            return {TextPr: inheritedTextPr};
        },
        GetParent: function() {
            return paraParent;
        },
        IsParaPrCompiled: function() {
            return true;
        },
        Get_Theme: function() {
            return null;
        },
        Get_ColorMap: function() {
            return null;
        },
        bFromDocument: true,
        IsInFixedForm: function() {
            return false;
        },
        getLayoutFontSizeCoefficient: function() {
            return 1;
        },
        Style_Get: function() {
            return undefined;
        }
    };

    const isMathRun = !!options.isMathRun;
    const run = new ctx.ParaRun(paragraph, isMathRun);

    if (options.rStyle !== undefined || options.styleEastAsia !== undefined) {
        run.Pr.RStyle = options.rStyle || 'Style1';
    }
    tag(run.Pr, 'direct');
    if (options.directEastAsia !== undefined) {
        run.Pr.Lang = Object.assign(
            {},
            run.Pr.Lang,
            {EastAsia: options.directEastAsia}
        );
    }

    if (options.isStyleHyperlink !== undefined) {
        run.IsStyleHyperlink = function() {
            return options.isStyleHyperlink;
        };
    }
    if (options.isInHyperlinkInTOC !== undefined) {
        run.IsInHyperlinkInTOC = function() {
            return options.isInHyperlinkInTOC;
        };
    }

    if (isMathRun) {
        if (options.hasParent !== false) {
            const ctrPrp = tag(new ctx.CTextPr(), 'ctrPrp');
            if (options.ctrPrpEastAsia !== undefined) {
                ctrPrp.Lang = {EastAsia: options.ctrPrpEastAsia};
            }
            run.Parent = {
                GetCtrPrp: function() {
                    return ctrPrp;
                }
            };
        }

        if (options.isPlaceholder) {
            run.Content = [{
                IsPlaceholder: function() {
                    return true;
                }
            }];
        }

        const mathTextPr = tag({Lang: {}}, 'mathText');
        if (options.mathTextEastAsia !== undefined) {
            mathTextPr.Lang = {EastAsia: options.mathTextEastAsia};
        }
        run.MathPrp = {
            GetTxtPrp: function() {
                return mathTextPr;
            },
            GetCompiled_ScrStyles: function() {
                return {nor: options.isNormalText !== false};
            }
        };
    }

    return {run: run, paragraph: paragraph, inheritedTextPr: inheritedTextPr};
}

function createRunCompileFixture(options) {
    options = options || {};
    const ctx = options.harness || createRunHarnessContext();
    const built = buildRun(ctx, options);

    return {
        run: built.run,
        cache: ctx.cache,
        mergeLog: ctx.mergeLog,
        textPr: ctx.CTextPr,
        harness: ctx,
        paragraph: built.paragraph,
        inheritedTextPr: built.inheritedTextPr
    };
}

test('run snapshot captures style before direct language', function() {
    const fixture = createRunCompileFixture({
        paragraphEastAsia: 1033,
        styleEastAsia: 2052,
        directEastAsia: 3084
    });

    const finalTextPr = fixture.run.Internal_Compile_Pr();
    const firstMergeOrder = fixture.mergeLog.slice(0, 3);

    assert.equal(finalTextPr.Lang.EastAsia, 3084);
    assert.equal(
        typeof fixture.run.GetCompiledEastAsiaBeforeDirect,
        'function'
    );
    assert.equal(fixture.run.GetCompiledEastAsiaBeforeDirect(), 2052);
    assert.deepEqual(firstMergeOrder, ['paragraph', 'style', 'direct']);
});

test('character style value remains when direct language is absent',
function() {
    const fixture = createRunCompileFixture({
        paragraphEastAsia: 1033,
        styleEastAsia: 3084
    });

    const finalTextPr = fixture.run.Internal_Compile_Pr();

    assert.equal(finalTextPr.Lang.EastAsia, 3084);
    assert.equal(fixture.run.GetCompiledEastAsiaBeforeDirect(), 3084);
});

test('based-on Korean character style remains compiled', function() {
    const fixture = createRunCompileFixture({
        paragraphEastAsia: 1033,
        styleEastAsia: 1042
    });

    const finalTextPr = fixture.run.Internal_Compile_Pr();

    assert.equal(finalTextPr.Lang.EastAsia, 1042);
    assert.equal(fixture.run.GetCompiledEastAsiaBeforeDirect(), 1042);
});

test('Hyperlink in TOC keeps its existing style exception', function() {
    const fixture = createRunCompileFixture({
        paragraphEastAsia: 1033,
        styleEastAsia: 2052,
        directEastAsia: 3084,
        isStyleHyperlink: true,
        isInHyperlinkInTOC: true
    });

    const finalTextPr = fixture.run.Internal_Compile_Pr();
    const firstMergeOrder = fixture.mergeLog.slice(0, 2);

    assert.equal(finalTextPr.Lang.EastAsia, 3084);
    assert.deepEqual(firstMergeOrder, ['paragraph', 'direct']);
    assert.equal(fixture.mergeLog.indexOf('style'), -1);
    assert.equal(fixture.run.GetCompiledEastAsiaBeforeDirect(), 1033);
});

test('math placeholder snapshots CtrPrp before direct properties',
function() {
    const fixture = createRunCompileFixture({
        isMathRun: true,
        isPlaceholder: true,
        paragraphEastAsia: 1033,
        ctrPrpEastAsia: 2052,
        directEastAsia: 3084
    });

    const finalTextPr = fixture.run.Internal_Compile_Pr();
    const firstMergeOrder = fixture.mergeLog.slice(0, 3);

    assert.equal(finalTextPr.Lang.EastAsia, 3084);
    assert.deepEqual(firstMergeOrder, ['paragraph', 'ctrPrp', 'direct']);
    assert.equal(fixture.run.GetCompiledEastAsiaBeforeDirect(), 2052);
});

test('math text properties remain after direct merge', function() {
    const fixture = createRunCompileFixture({
        isMathRun: true,
        isPlaceholder: false,
        isNormalText: false,
        styles: null,
        paragraphEastAsia: 1033,
        directEastAsia: 3084,
        mathTextEastAsia: 5124
    });

    const finalTextPr = fixture.run.Internal_Compile_Pr();
    const firstMergeOrder = fixture.mergeLog.slice(0, 3);

    assert.equal(finalTextPr.Lang.EastAsia, 5124);
    assert.deepEqual(firstMergeOrder, ['paragraph', 'direct', 'mathText']);
});

test('early returns clear an older snapshot', function() {
    const missingParagraph = createRunCompileFixture({
        paragraphMissing: true
    });
    missingParagraph.run.CompiledEastAsiaBeforeDirect = 2052;
    const defaultTextPr1 = missingParagraph.run.Internal_Compile_Pr();

    assert.equal(defaultTextPr1.Lang.EastAsia, 1033);
    assert.equal(missingParagraph.run.CompiledEastAsiaBeforeDirect, 1033);

    const missingMathParent = createRunCompileFixture({
        isMathRun: true,
        hasParent: false,
        paragraphEastAsia: 1033
    });
    missingMathParent.run.CompiledEastAsiaBeforeDirect = 2052;
    const defaultTextPr2 = missingMathParent.run.Internal_Compile_Pr();

    assert.equal(defaultTextPr2.Lang.EastAsia, 1033);
    assert.equal(missingMathParent.run.CompiledEastAsiaBeforeDirect, 1033);
});

test('recompilation refreshes the run-owned snapshot', function() {
    const fixture = createRunCompileFixture({paragraphEastAsia: 2052});

    assert.equal(fixture.run.GetCompiledEastAsiaBeforeDirect(), 2052);

    fixture.inheritedTextPr.Lang = {EastAsia: 1042};
    fixture.run.RecalcInfo.TextPr = true;

    assert.equal(fixture.run.GetCompiledEastAsiaBeforeDirect(), 1042);
});

test('deduplicated final properties do not share snapshots', function() {
    const first = createRunCompileFixture({
        paragraphEastAsia: 2052,
        directEastAsia: 3084
    });
    const second = createRunCompileFixture({
        harness: first.harness,
        paragraphEastAsia: 1042,
        directEastAsia: 3084
    });

    const firstSnapshot = first.run.GetCompiledEastAsiaBeforeDirect();
    const secondSnapshot = second.run.GetCompiledEastAsiaBeforeDirect();

    assert.equal(firstSnapshot, 2052);
    assert.equal(secondSnapshot, 1042);
    assert.equal(first.run.CompiledPr.Lang.EastAsia, 3084);
    assert.equal(first.run.CompiledPr, second.run.CompiledPr);
});
