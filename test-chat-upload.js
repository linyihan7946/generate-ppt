"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var axios_1 = __importDefault(require("axios"));
var form_data_1 = __importDefault(require("form-data"));
var fs_1 = __importDefault(require("fs"));
var path_1 = __importDefault(require("path"));
var BASE_URL = 'http://localhost:3000';
function testChatWithFile() {
    return __awaiter(this, void 0, void 0, function () {
        var response, err_1, response, err_2, formData, docxPath, response, err_3, formData, mdPath, response, err_4, formData, docxPath, response, outputPath, err_5;
        var _a, _b, _c, _d, _e, _f, _g;
        return __generator(this, function (_h) {
            switch (_h.label) {
                case 0:
                    console.log('=== Test 1: Chat API Health Check ===');
                    _h.label = 1;
                case 1:
                    _h.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, axios_1.default.get(BASE_URL)];
                case 2:
                    response = _h.sent();
                    console.log('✅ Server is running:', response.status === 200);
                    return [3 /*break*/, 4];
                case 3:
                    err_1 = _h.sent();
                    console.error('❌ Server is not running:', err_1.message);
                    console.log('Please start the server with: npm start');
                    return [2 /*return*/];
                case 4:
                    console.log('\n=== Test 2: Chat with text message ===');
                    _h.label = 5;
                case 5:
                    _h.trys.push([5, 7, , 8]);
                    return [4 /*yield*/, axios_1.default.post("".concat(BASE_URL, "/api/chat"), {
                            messages: [
                                { role: 'user', content: '我想做一个关于AI技术的工作汇报PPT' }
                            ]
                        }, {
                            headers: { 'Content-Type': 'application/json' }
                        })];
                case 6:
                    response = _h.sent();
                    console.log('✅ Response received:', response.data.reply ? 'YES' : 'NO');
                    console.log('Reply preview:', ((_a = response.data.reply) === null || _a === void 0 ? void 0 : _a.substring(0, 100)) + '...');
                    return [3 /*break*/, 8];
                case 7:
                    err_2 = _h.sent();
                    console.error('❌ Chat with text failed:', ((_b = err_2.response) === null || _b === void 0 ? void 0 : _b.data) || err_2.message);
                    return [3 /*break*/, 8];
                case 8:
                    console.log('\n=== Test 3: Chat with file upload (DOCX) ===');
                    _h.label = 9;
                case 9:
                    _h.trys.push([9, 13, , 14]);
                    formData = new form_data_1.default();
                    formData.append('text', '请根据这份文档生成PPT');
                    formData.append('messages', JSON.stringify([]));
                    docxPath = path_1.default.join(__dirname, 'input', '计算机发展史.docx');
                    if (!fs_1.default.existsSync(docxPath)) return [3 /*break*/, 11];
                    formData.append('files', fs_1.default.createReadStream(docxPath), {
                        filename: '计算机发展史.docx',
                        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    });
                    return [4 /*yield*/, axios_1.default.post("".concat(BASE_URL, "/api/chat"), formData, {
                            headers: formData.getHeaders()
                        })];
                case 10:
                    response = _h.sent();
                    console.log('✅ File uploaded successfully');
                    console.log('Reply:', ((_c = response.data.reply) === null || _c === void 0 ? void 0 : _c.substring(0, 150)) + '...');
                    console.log('Download URL:', response.data.downloadUrl || 'NOT_PROVIDED');
                    return [3 /*break*/, 12];
                case 11:
                    console.log('⚠️ Test file not found:', docxPath);
                    _h.label = 12;
                case 12: return [3 /*break*/, 14];
                case 13:
                    err_3 = _h.sent();
                    console.error('❌ Chat with file failed:', ((_d = err_3.response) === null || _d === void 0 ? void 0 : _d.data) || err_3.message);
                    return [3 /*break*/, 14];
                case 14:
                    console.log('\n=== Test 4: Chat with Markdown file ===');
                    _h.label = 15;
                case 15:
                    _h.trys.push([15, 19, , 20]);
                    formData = new form_data_1.default();
                    formData.append('text', '请根据这份文档生成PPT');
                    formData.append('messages', JSON.stringify([]));
                    mdPath = path_1.default.join(__dirname, 'test.md');
                    if (!fs_1.default.existsSync(mdPath)) return [3 /*break*/, 17];
                    formData.append('files', fs_1.default.createReadStream(mdPath), {
                        filename: 'test.md',
                        contentType: 'text/markdown'
                    });
                    return [4 /*yield*/, axios_1.default.post("".concat(BASE_URL, "/api/chat"), formData, {
                            headers: formData.getHeaders()
                        })];
                case 16:
                    response = _h.sent();
                    console.log('✅ Markdown file uploaded successfully');
                    console.log('Reply:', ((_e = response.data.reply) === null || _e === void 0 ? void 0 : _e.substring(0, 150)) + '...');
                    console.log('Download URL:', response.data.downloadUrl || 'NOT_PROVIDED');
                    return [3 /*break*/, 18];
                case 17:
                    console.log('⚠️ Test file not found:', mdPath);
                    _h.label = 18;
                case 18: return [3 /*break*/, 20];
                case 19:
                    err_4 = _h.sent();
                    console.error('❌ Chat with Markdown file failed:', ((_f = err_4.response) === null || _f === void 0 ? void 0 : _f.data) || err_4.message);
                    return [3 /*break*/, 20];
                case 20:
                    console.log('\n=== Test 5: Generate PPT endpoint (direct file upload) ===');
                    _h.label = 21;
                case 21:
                    _h.trys.push([21, 24, , 25]);
                    formData = new form_data_1.default();
                    docxPath = path_1.default.join(__dirname, 'input', '计算机发展史.docx');
                    if (!fs_1.default.existsSync(docxPath)) return [3 /*break*/, 23];
                    formData.append('file', fs_1.default.createReadStream(docxPath), {
                        filename: '计算机发展史.docx',
                        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    });
                    formData.append('plannerMode', 'creative');
                    return [4 /*yield*/, axios_1.default.post("".concat(BASE_URL, "/generate-ppt"), formData, {
                            headers: formData.getHeaders(),
                            responseType: 'arraybuffer'
                        })];
                case 22:
                    response = _h.sent();
                    console.log('✅ PPT generated successfully');
                    console.log('Content-Type:', response.headers['content-type']);
                    console.log('Content-Length:', response.data.byteLength, 'bytes');
                    outputPath = path_1.default.join(__dirname, 'test-output.pptx');
                    fs_1.default.writeFileSync(outputPath, Buffer.from(response.data));
                    console.log('✅ PPT saved to:', outputPath);
                    _h.label = 23;
                case 23: return [3 /*break*/, 25];
                case 24:
                    err_5 = _h.sent();
                    console.error('❌ Generate PPT failed:', ((_g = err_5.response) === null || _g === void 0 ? void 0 : _g.data) || err_5.message);
                    return [3 /*break*/, 25];
                case 25:
                    console.log('\n=== All tests completed ===');
                    return [2 /*return*/];
            }
        });
    });
}
testChatWithFile();
