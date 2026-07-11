'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SDK_ROOT = path.resolve(__dirname, '../../..');

function createLooseStub() {
    let proxy;
    const target = function() {
        return proxy;
    };

    proxy = new Proxy(target, {
        get: function(_target, property) {
            if (property === Symbol.toPrimitive) {
                return function() {
                    return 0;
                };
            }
            if (property === 'prototype') {
                return {};
            }
            return proxy;
        },
        apply: function() {
            return proxy;
        },
        construct: function() {
            return proxy;
        }
    });

    return proxy;
}

// Under `looseGlobals`, the sandbox is wrapped in a Proxy whose `has` trap
// unconditionally returns true (see below) so that any unresolved bare
// identifier in legacy source resolves to a harmless callable/constructible
// stub instead of throwing a ReferenceError. That trap makes the sandbox
// *claim* ownership of every property name to the vm's identifier
// resolution, which means real language intrinsics (Object, Array, Math,
// JSON, undefined, ...) must be listed explicitly as own properties of the
// sandbox too - otherwise a bare `Object.create(...)` inside loaded source
// would silently resolve `Object` itself to the stub instead of the real
// constructor. These are Node's real intrinsics, reused across the realm
// boundary; that is safe for the constructor/method calls this harness
// exercises.
const NATIVE_INTRINSICS = {
    Object: Object,
    Array: Array,
    Math: Math,
    JSON: JSON,
    Date: Date,
    RegExp: RegExp,
    Map: Map,
    Set: Set,
    WeakMap: WeakMap,
    WeakSet: WeakSet,
    Promise: Promise,
    Symbol: Symbol,
    Error: Error,
    TypeError: TypeError,
    RangeError: RangeError,
    SyntaxError: SyntaxError,
    Number: Number,
    String: String,
    Boolean: Boolean,
    Function: Function,
    isNaN: isNaN,
    isFinite: isFinite,
    parseInt: parseInt,
    parseFloat: parseFloat,
    NaN: NaN,
    Infinity: Infinity,
    undefined: undefined
};

function createSdkHarness(options) {
    options = options || {};

    const base = Object.assign({
        console: console,
        Uint8Array: Uint8Array,
        Asc: {},
        AscBidi: {},
        AscCommon: {},
        AscWord: {},
        AscFonts: {
            allocate: function(size) {
                return new Uint8Array(size);
            }
        }
    }, options.looseGlobals ? NATIVE_INTRINSICS : {}, options.globals || {});
    const looseStub = createLooseStub();
    const sandbox = options.looseGlobals
        ? new Proxy(base, {
            has: function() {
                return true;
            },
            get: function(target, property) {
                if (property === Symbol.unscopables) {
                    return undefined;
                }
                return Object.prototype.hasOwnProperty.call(
                    target,
                    property
                ) ? target[property] : looseStub;
            }
        })
        : base;

    sandbox.window = sandbox;
    sandbox.globalThis = sandbox;

    const context = vm.createContext(sandbox);

    function load(relativePath) {
        const filename = path.resolve(SDK_ROOT, relativePath);
        const rootPrefix = SDK_ROOT + path.sep;
        if (!filename.startsWith(rootPrefix)) {
            throw new Error('Source escapes SDK root: ' + relativePath);
        }

        const source = fs.readFileSync(filename, 'utf8');
        vm.runInContext(source, context, {filename: filename});
        return sandbox;
    }

    function loadResolverStack() {
        load('common/commonDefines.js');
        load('common/bidi/bidi-types.js');
        load('word/Editor/Paragraph/Run/FontClassification.js');
        load('word/Editor/Paragraph/Run/LanguageResolver.js');
        return sandbox;
    }

    return {
        context: context,
        load: load,
        loadResolverStack: loadResolverStack,
        sandbox: sandbox
    };
}

function createTextPr(lang) {
    return {
        Bold: false,
        BoldCS: false,
        Caps: false,
        FontSize: 11,
        FontSizeCS: 11,
        Italic: false,
        ItalicCS: false,
        Lang: Object.assign({}, lang),
        RFonts: {
            Ascii: {Name: 'Arial'},
            CS: {Name: 'Arial'},
            EastAsia: {Name: 'SimSun'},
            HAnsi: {Name: 'Arial'}
        },
        ReplaceThemeFonts: function() {}
    };
}

function createTextElement(options) {
    options = options || {};
    const codePoint = options.codePoint === undefined
        ? 0x61
        : options.codePoint;

    return {
        GetCharForSpellCheck: function() {
            return String.fromCodePoint(codePoint);
        },
        GetCodePoint: function() {
            return codePoint;
        },
        GetDirectionFlag: function() {
            return options.direction;
        },
        GetFontSlot: function() {
            return options.fontSlot;
        },
        IsDot: function() {
            return false;
        },
        IsInstrText: function() {
            return !!options.isInstruction;
        },
        IsPunctuation: function() {
            return !!options.isPunctuation;
        },
        IsText: function() {
            return options.isText !== false;
        }
    };
}

module.exports = {
    SDK_ROOT: SDK_ROOT,
    createSdkHarness: createSdkHarness,
    createTextElement: createTextElement,
    createTextPr: createTextPr
};
