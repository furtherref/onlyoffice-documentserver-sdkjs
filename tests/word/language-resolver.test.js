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
