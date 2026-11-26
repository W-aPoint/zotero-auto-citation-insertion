/* A local, narrowly scoped Word client for the installed ODF/DOCX Scan converter. */
var route = "/word-zotero-scan/v1";
var secret, converterScope;
var ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function nativeConverter() {
    if (converterScope) return converterScope;
    let scope = { Zotero };
    // Upstream also exports CitationUtils on the scanner window.
    scope.window = scope;
    try {
        Services.scriptloader.loadSubScript("chrome://odf-scan/content/citationUtils.js", scope);
        Services.scriptloader.loadSubScript("chrome://odf-scan/content/docxConverter.js", scope);
        if (typeof scope.DOCXConverter.markersToCitations !== "function") throw Error();
    } catch (e) {
        throw Error("无法加载 ODF/DOCX Scan 的 DOCX 转换器：" + (e.message || e));
    }
    converterScope = scope;
    return scope;
}

function writableCollection(target) {
    const col = Zotero.Collections.get(Number(target));
    if (!col || col.deleted || !col.library.editable) throw Error("这个分类已移除或不可写，请重新选择。");
    return col;
}

function listCollections() {
    const output = [];
    for (const lib of Zotero.Libraries.getAll()) {
        if (!lib.editable) continue;
        const collections = Zotero.Collections.getByLibrary(lib.libraryID, true);
        for (const col of collections) {
            if (!col.deleted) output.push({
                id: col.id, key: col.key, parent: col.parentID || 0, name: col.name,
                library: lib.libraryID, libraryName: lib.name
            });
        }
    }
    return output;
}

async function resolveMarker(code) {
    const m = /^(zu|zg):(\d+):([A-Z0-9]{8})$/.exec(code);
    if (!m) throw Error("无效的 Scannable Cite 条目标识：" + code);
    let lib;
    if (m[1] === "zu") {
        if (m[2] !== "0" && String(Zotero.Users.getCurrentUserID()) !== m[2])
            throw Error("标记属于另一个 Zotero 个人文库：" + code);
        lib = Zotero.Libraries.userLibraryID;
    } else {
        lib = Zotero.Groups.getLibraryIDFromGroupID(Number(m[2]));
    }
    const item = lib && await Zotero.Items.getByLibraryAndKeyAsync(lib, m[3]);
    if (!item || item.deleted || !item.isRegularItem()) throw Error("本机 Zotero 中找不到文献：" + code);
    return item;
}

async function resolveCitation(uris) {
    if (!Array.isArray(uris)) throw Error("已有引文缺少 Zotero 文献 URI。");
    for (let uri of uris) {
        if (typeof uri !== "string") continue;
        try {
            const item = /^(zu|zg):/.test(uri) ? await resolveMarker(uri) : await Zotero.URI.getURIItem(uri);
            if (item && !item.deleted && item.isRegularItem()) return item;
            // The official Word integration can retain a replaced item's URI after a merge.
            if (/^https?:\/\/(?:www\.)?zotero\.org\/(?:users|groups)\//.test(uri)) {
                const merged = await Zotero.Relations.getByPredicateAndObject("item", "dc:replaces", uri);
                const live = merged.filter(it => !it.deleted && it.isRegularItem());
                if (live.length === 1) return live[0];
            }
        } catch (_) { /* Check the next alias stored by Zotero. */ }
    }
    throw Error("已有引文对应的文献不在当前文库：" + uris.join(" / "));
}

function inLibrary(item, col) {
    if (item.libraryID !== col.libraryID)
        throw Error("文献与所选分类属于不同文库：" + item.getField("title"));
    return item;
}

async function prepare(data) {
    const col = writableCollection(data.target);
    if (!Array.isArray(data.groups) || !Array.isArray(data.existing)) throw Error("扫描请求无效。");
    const items = new Map(), fields = [];
    const scope = data.groups.length ? nativeConverter() : null;
    for (const group of data.groups) {
        const matches = Array.from(group.matchAll(/\{([^{}|]*)\|([^{}|]*)\|([^{}|]*)\|([^{}|]*)\|([^{}|]*)\}/g));
        if (!matches.length || matches.map(m => m[0]).join("") !== group) throw Error("发现无效的五字段标记：" + group);
        // Feed the converter one citation group at a time. It strips XML tags in
        // marker text, so shield literal '<' until the upstream field is complete.
        let token;
        do { token = "LT" + Services.uuid.generateUUID().toString().replace(/[^a-z0-9]/gi, ""); } while (group.includes(token));
        let input = "";
        const expected = [];
        for (const m of matches) {
            const item = inLibrary(await resolveMarker(m[5].trim()), col);
            const uri = Zotero.URI.getItemURI(item);
            items.set(item.id, uri);
            expected.push(item.id);
            input += "{" + [m[1], m[2], m[3], m[4], uri].join("|").replace(/</g, token) + "}";
        }
        const xml = '<w:p xmlns:w="' + ns + '"><w:r><w:t>' + input + '</w:t></w:r></w:p>';
        const converted = (await scope.DOCXConverter.markersToCitations(xml)).split(token).join("&lt;");
        const win = Zotero.getMainWindow();
        const dom = new win.DOMParser().parseFromString(converted, "application/xml");
        const instructions = dom.getElementsByTagNameNS(ns,"instrText");
        if (instructions.length !== 1) throw Error("ODF/DOCX Scan 未生成预期的组合引文。");
        const instruction = instructions[0].textContent.trim();
        const prefix = "ADDIN ZOTERO_ITEM CSL_CITATION ";
        if (!instruction.startsWith(prefix)) throw Error("ODF/DOCX Scan 返回的引文字段无法识别。");
        const cite = JSON.parse(instruction.slice(prefix.length));
        if (cite.citationItems.length !== expected.length) throw Error("转换前后条目数不同。");
        for (let i = 0; i < expected.length; i++) {
            if ((await resolveCitation(cite.citationItems[i].uris)).id !== expected[i]) throw Error("转换前后文献不一致。");
        }
        fields.push({ instruction, display: cite.properties.plainCitation });
    }
    for (const uris of data.existing) {
        const item = inLibrary(await resolveCitation(uris), col);
        items.set(item.id, Zotero.URI.getItemURI(item));
    }
    return { fields, uris: Array.from(items.values()) };
}

async function archive(data) {
    const col = writableCollection(data.target);
    if (!Array.isArray(data.uris)) throw Error("归档请求无效。");
    const items = new Map();
    for (const uri of data.uris) {
        const item = inLibrary(await resolveCitation([uri]), col);
        items.set(item.id, item);
    }
    let added = 0;
    await Zotero.DB.executeTransaction(async () => {
        for (const item of items.values()) {
            if (item.inCollection(col.id)) continue;
            item.addToCollection(col.id);
            await item.save();
            added++;
        }
    });
    return { total: items.size, added };
}

var Endpoint = function () {};
Endpoint.prototype = {
    supportedMethods: ["POST"],
    supportedDataTypes: ["application/json"],
    permitBookmarklet: false,
    async init({data}) {
        const reply = (code, value) => [code, "application/json", JSON.stringify(value)];
        if (!secret || !data || data.auth !== secret) return reply(403, {error:"Word 与 Zotero 的本机连接未获授权。"});
        try {
            if (data.action === "collections") return reply(200, {collections: listCollections()});
            if (data.action === "prepare") return reply(200, await prepare(data));
            if (data.action === "archive") return reply(200, await archive(data));
            return reply(400, {error:"未知操作。"});
        } catch (e) {
            return reply(422, {error:String(e.message || e)});
        }
    }
};
function install() {}
function uninstall() {}
async function startup() {
    await Zotero.initializationPromise;
    const path = PathUtils.join(Services.env.get("LOCALAPPDATA"), "WordZoteroScan", "connection.key");
    secret = (await IOUtils.readUTF8(path)).trim();
    if (!/^[a-f0-9]{64}$/.test(secret)) throw Error("Word Zotero Scan connection key is missing.");
    Zotero.Server.Endpoints[route] = Endpoint;
}
function shutdown() {
    if (Zotero.Server.Endpoints[route] === Endpoint) delete Zotero.Server.Endpoints[route];
    secret = null;
    converterScope = null;
}
