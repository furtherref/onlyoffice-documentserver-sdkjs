/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function(window)
{
    const oAsc                 = window["Asc"] || {};
    const oNameToLcid          = oAsc.g_oLcidNameToIdMap || {};
    const arrIdeographLanguage = oAsc.availableIdeographLanguages || [];
    const oIdeographLcids      = Object.create(null);

    function AddIdeographLanguage(sName)
    {
        if (!sName
            || !Object.prototype.hasOwnProperty.call(oNameToLcid, sName))
            return;

        oIdeographLcids[oNameToLcid[sName]] = true;
    }

    for (let nIndex = 0; nIndex < arrIdeographLanguage.length; ++nIndex)
    {
        let sName = arrIdeographLanguage[nIndex];
        AddIdeographLanguage(sName);

        let nSeparator = sName.indexOf("-");
        if (nSeparator > 0)
            AddIdeographLanguage(sName.slice(0, nSeparator));
    }

    function IsIdeographLanguage(nLcid)
    {
        return true === oIdeographLcids[nLcid];
    }

    function ResolveTextLanguage(
        oLang,
        nFontSlot,
        nDirectionFlag,
        nEastAsiaBeforeDirect)
    {
        if (!oLang)
            return undefined;

        if (AscBidi.DIRECTION_FLAG.RTL === nDirectionFlag
            || 0 !== (nFontSlot & AscWord.fontslot_CS))
        {
            return oLang.Bidi;
        }

        if (AscBidi.DIRECTION_FLAG.LTR !== nDirectionFlag)
            return undefined;

        if (0 !== (nFontSlot & AscWord.fontslot_EastAsia))
        {
            if (IsIdeographLanguage(oLang.EastAsia))
                return oLang.EastAsia;

            if (IsIdeographLanguage(nEastAsiaBeforeDirect))
                return nEastAsiaBeforeDirect;

            if (undefined !== oLang.EastAsia
                && null !== oLang.EastAsia)
            {
                return oLang.EastAsia;
            }

            return oLang.Val;
        }

        if (0 !== (nFontSlot & AscWord.fontslot_ASCII)
            || 0 !== (nFontSlot & AscWord.fontslot_HAnsi))
        {
            return oLang.Val;
        }

        return undefined;
    }

    function ResolveRunElementLanguage(
        oElement,
        oTextPr,
        nEastAsiaBeforeDirect)
    {
        if (!oElement
            || !oTextPr
            || !oTextPr.Lang
            || !oElement.IsText
            || !oElement.IsText())
        {
            return undefined;
        }

        let nDirectionFlag = oElement.GetDirectionFlag();
        if (AscBidi.DIRECTION_FLAG.LTR !== nDirectionFlag
            && AscBidi.DIRECTION_FLAG.RTL !== nDirectionFlag)
        {
            return undefined;
        }

        return ResolveTextLanguage(
            oTextPr.Lang,
            oElement.GetFontSlot(oTextPr),
            nDirectionFlag,
            nEastAsiaBeforeDirect
        );
    }

    window["AscWord"] = window["AscWord"] || {};
    window["AscWord"].IsIdeographLanguage = IsIdeographLanguage;
    window["AscWord"].ResolveTextLanguage = ResolveTextLanguage;
    window["AscWord"].ResolveRunElementLanguage =
        ResolveRunElementLanguage;

})(window);
