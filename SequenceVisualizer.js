// @ts-nocheck
/** @OnlyCurrentDoc */
function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Edit Distance')
        .addItem("Calculate edit distance", 'promptXY')
        .addSeparator()
        .addItem("Correct word strictly", 'correctWordOptimized')
        .addSeparator()
        .addItem("Correct word", 'correctWordVerbose')
        .addToUi();
}

//add edit distance to result
function promptXY() {
    let ui = SpreadsheetApp.getUi();

    let promptX = ui.prompt(
        'Populate the field',
        'String X',
        ui.ButtonSet.OK_CANCEL);

    const textX = promptX.getResponseText();

    if (textX == null || textX.length == 0) return;

    let promptY = ui.prompt(
        'Populate the field',
        'String Y',
        ui.ButtonSet.OK_CANCEL);

    const textY = promptY.getResponseText();

    if (textY == null || textY.length == 0) return;

    let promptGap = ui.prompt(
        'Populate the field',
        'Gap Cost',
        ui.ButtonSet.OK_CANCEL);

    const gapCost = Math.round(Number(promptGap.getResponseText()));

    if (gapCost == null || gapCost.length == 0) return;

    const opt = editDistance(textX, textY, gapCost); //I need typescript to see where my issue is with repeating decimals causing issues, no decimals for now... ;(
    displayOPT(textX, textY, gapCost, opt);
}

//cost of swapping a & b where a == b
const EQUAL_CHAR_SWAP_COST = 0.0;

function getSwapCost(a, b, ccm) {
    a = a.toLowerCase();
    b = b.toLowerCase();
    if (a == b) return EQUAL_CHAR_SWAP_COST; //unneeded but still keeping it in in-case a future ccm has a diff result
    let p1 = ccm[a];
    let p2 = ccm[b];
    let xDiff = Math.pow(p1.x - p2.x, 2);
    let yDiff = Math.pow(p1.y - p2.y, 2);
    let res = Math.sqrt(xDiff + yDiff);
    return res;
}

function getSwapCostOptimized(a, b, ccm) {
    a = a.toLowerCase();
    b = b.toLowerCase();
    if (a == b) return EQUAL_CHAR_SWAP_COST; //unneeded but still keeping it in in-case a future ccm has a diff result
    let res = ccm[Number(a.charCodeAt(0) - 'a'.charCodeAt(0))][Number(b.charCodeAt(0) - 'a'.charCodeAt(0))];
    return Math.round(res);
}

const ccm = get2dAlphaCCM();

//cost of turning X into Y given gapCost
function editDistance(X, Y, gapCost) {
    let A = get2DArray(X.length + 1, Y.length + 1);
    let path = get2DArray(X.length + 1, Y.length + 1); //map of {i,j} -> parent cell {i, j} of calculation
    let operations = get2DArray(X.length + 1, Y.length + 1);
    
    //init values for when Y is empty
    for (let i = X.length; i >= 0; i--) {
        A[i][0] = gapCost * Number(i);
        if (i != 0) {
            path[i][0] = { y: i - 1, x: 0, dir: 1 };
            operations[i][0] = "Gap";
        }
    }

    //init values for when X is empty
    for (let j = Y.length; j >= 0; j--) {
        A[0][j] = gapCost * Number(j);
        if (j != 0) {
            path[0][j] = { y: 0, x: j - 1, dir: 2 };
            operations[0][j] = "Gap";
        }
    }

    let m = X.length;
    let n = Y.length;

    for (let j = 0; j < n; j++) {
        for (let i = 0; i < m; i++) {
            const a = X.charAt(i);
            const b = Y.charAt(j);
            const swapCost = getSwapCostOptimized(a, b, ccm);

            const swapPath = swapCost + Number(A[i][j]);
            const gapPathA = gapCost + Number(A[i][j + 1]);
            const gapPathB = gapCost + Number(A[i + 1][j]);

            A[i + 1][j + 1] = Math.min(swapPath,
                Math.min(gapPathA, gapPathB));

            if (A[i + 1][j + 1] == swapPath) {
                path[i + 1][j + 1] = { y: i, x: j, dir: 0 };
                operations[i + 1][j + 1] = "Swap";
                if (swapPath == gapPathA || swapPath == gapPathB) {
                  operations[i + 1][j + 1] += " | Gap";
                }
            } else if (A[i + 1][j + 1] == gapPathA) {
                path[i + 1][j + 1] = { y: i, x: j + 1, dir: 1 };
                operations[i + 1][j + 1] = "Gap";
            } else {
                path[i + 1][j + 1] = { y: i + 1, x: j, dir: 2 };
                operations[i + 1][j + 1] = "Gap";
            }
        }
    }
    return [A, A[m][n], path, operations];
}

function getArrowFromDir(dir) {
    if (dir == 0) return "↗";
    if (dir == 1) return "↑";
    if (dir == 2) return "→";
    return "";
}

function getColor(depth, maxDepth) {
    let r = 0;
    let g = 255 * (depth / maxDepth);
    let b = 0;
    return [r, g, b];
}

function isColorSimilar(c1, c2, threshold) {
    let r1 = c1[0];
    let g1 = c1[1];
    let b1 = c1[2];

    let r2 = c2[0];
    let g2 = c2[1];
    let b2 = c2[2];

    let r = (Math.abs(r1 - r2) / 255) * 100;
    let g = (Math.abs(g1 - g2) / 255) * 100;
    let b = (Math.abs(b1 - b2) / 255) * 100;

    return ((r + g + b) / 3) <= threshold;
}

function setShortestPathParents(arr, path, operations, sheet) {
    //dir mapping: 0 -> (i - 1, j - 1), ↗
    //dir mapping: 1 -> (i -1, j), ↑
    //dir mapping: 2 -> (i, j - 1), →
    let shortestPath = [];
    let i = arr.length - 1;
    let j = arr[i].length - 1;
    let stack = [];
    stack.push({ i: i, j: j });
    while (stack.length > 0) {
        let curr = stack.pop();
        let currI = curr.i;
        let currJ = curr.j;
        shortestPath.push({ i: currI, j: currJ });//keep track so we can color with proper depth
        let currentPath = path[currI][currJ];
        if (currentPath == null || currentPath.x == null) continue;
        arr[currI][currJ] += "\n" + operations[currI][currJ];
        arr[currentPath.y][currentPath.x] = "\t\t" + getArrowFromDir(currentPath.dir) + "\n" + arr[currentPath.y][currentPath.x];
        //currentPath.x + 2 because sheets 0,0 is 1,1 so our i's are fine since we are using 1-indexed arr.length - currentPath.y but j needs to add 2 which is effectively adding 1 as our offset from the column for the characters of the word
        stack.push({ i: currentPath.y, j: currentPath.x });
    }
    let lengthOfPath = shortestPath.length;
    if (lengthOfPath <= 0) return arr;
    for (let p = shortestPath.length - 1; p >= 0; p--) {
        const currentPath = shortestPath[p];
        const color = getColor(shortestPath.length - p, lengthOfPath);
        const r = color[0];
        const g = color[1];
        const b = color[2];
        let currentCell = sheet.getRange(arr.length - currentPath.i, currentPath.j + 2, 1, 1);
        currentCell.setBackgroundRGB(r, g, b);
        if (isColorSimilar(color, [0, 0, 0], 20)) {
            currentCell.setFontColor("white");
        }
    }
    return arr;
}

function getEdits(X, Y, operations) {
    return ["", ""];
}

function displayOPT(X, Y, gapCost, opt) {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    let title = "Edit-Distance OPT(" + X + ", " + Y + ", " + gapCost + ")";
    var sheet = active.getSheetByName(title);

    if (sheet != null) {
        active.deleteSheet(sheet);
    }

    sheet = active.insertSheet();
    sheet.setName(title);

    if (opt == null) return;

    let xChars = Array.from(X);
    let yChars = Array.from(Y);

    let matrix = opt[0];
    let cost = opt[1];
    let path = opt[2];
    let operations = opt[3];

    matrix = setShortestPathParents(matrix, path, operations, sheet);

    for (let i = matrix.length - 1; i >= 0; i--) {
        let currentRow = matrix[i];
        currentRow.unshift(xChars[i - 1]);
        sheet.appendRow(currentRow);
    }

    //shift space from word column
    yChars.unshift(" ");
    //shift space from sentinel
    yChars.unshift(" ");
    sheet.appendRow(yChars);

    //Cell styling & formatting
    sheet.getRange(1, 1, X.length + 2, 1).setBackgroundRGB(189, 195, 199).setFontWeight("bold");
    sheet.getRange(X.length + 2, 1, 1, Y.length + 2).setBackgroundRGB(189, 195, 199).setFontWeight("bold");

    sheet.appendRow([" "]); //separator

    sheet.appendRow(["Edit Distance", cost]); //cost/result of editDistance between X && Y

    sheet.appendRow(["X", X]);
    sheet.appendRow(["Y", Y]);

    let edits = getEdits(X, Y, operations);
    sheet.appendRow(["Edited X", edits[0]]);
    sheet.appendRow(["Edited Y", edits[1]]);

    sheet.appendRow([" "]); //separator

    sheet.appendRow(["Gap Cost", gapCost])

    //Show swap costs
    sheet.appendRow(["Swap Costs"]);
    let compareLength = min(X.length, Y.length);
    for (let i = 0; i < compareLength; i++) {
        const a = X.charAt(i);
        const b = Y.charAt(i);
        const display = a + " ⇄ " + b + " = " + getSwapCostOptimized(a, b, ccm);
        sheet.appendRow([display]);
    }

    sheet.getRange(1, 1, (X.length + 10) * 2, (Y.length + 10) * 2).setHorizontalAlignment("center").setFontSize(12);
}

//Maybe use this to generate the 2d ccm array for speed purposes?
function getAlpaCCM() {
    const cartesian = {
        'q': {
            x: 0,
            y: 0
        },
        'w': {
            x: 1,
            y: 0
        },
        'e': {
            x: 2,
            y: 0
        },
        'r': {
            x: 3,
            y: 0
        },
        't': {
            x: 4,
            y: 0
        },
        'y': {
            x: 5,
            y: 0
        },
        'u': {
            x: 6,
            y: 0
        },
        'i': {
            x: 7,
            y: 0
        },
        'o': {
            x: 8,
            y: 0
        },
        'p': {
            x: 9,
            y: 0
        },
        'a': {
            x: 0,
            y: 1
        },
        's': {
            x: 1,
            y: 1
        },
        'd': {
            x: 2,
            y: 1
        },
        'f': {
            x: 3,
            y: 1
        },
        'g': {
            x: 4,
            y: 1
        },
        'h': {
            x: 5,
            y: 1
        },
        'j': {
            x: 6,
            y: 1
        },
        'k': {
            x: 7,
            y: 1
        },
        'l': {
            x: 8,
            y: 1
        },
        'z': {
            x: 0,
            y: 2
        },
        'x': {
            x: 1,
            y: 2
        },
        'c': {
            x: 2,
            y: 2
        },
        'v': {
            x: 3,
            y: 2
        },
        'b': {
            x: 4,
            y: 2
        },
        'n': {
            x: 5,
            y: 2
        },
        'm': {
            x: 6,
            y: 2
        },
    };
    return cartesian;
}

function get2dAlphaCCM() {
    let cartesian = getAlpaCCM();
    const ccm = get2DArray(26, 26);
    for (let i = 0; i < 26; i++) {
        let anchor = String.fromCharCode(i + 'a'.charCodeAt(0));
        for (let j = 0; j < 26; j++) {
            let compare = String.fromCharCode(j + 'a'.charCodeAt(0));
            ccm[i][j] = getSwapCost(anchor, compare, cartesian);
        }
    }
    return ccm;
}

function min(a, b) {
    if (a < b) return a;
    return b;
}

function get2DArray(rows, cols) {
    let result = []
    for (let i = 0; i < rows; i++) {
        result.push([])
        for (let j = 0; j < cols; j++) {
            result[i].push(0.0);
        }
    }
    return result;
}

function displaySummaryCorrectWord(optimized) {
    let ui = SpreadsheetApp.getUi();
    if (optimized) {
        ui.alert(
            'Correct word strictly',
            'This functionality is optimized by filtering out possible words that have that have different lengths, and different starting characters.',
            ui.ButtonSet.OK);
    } else {
        ui.alert(
            'Correct word',
            'This functionality compares with all words in the working dictionary.',
            ui.ButtonSet.OK);
    }
}

function correctWordOptimized() {
    displaySummaryCorrectWord(true);
    correctWord(true);
}

function correctWordVerbose() {
    displaySummaryCorrectWord(false);
    correctWord(false);
}

function correctWord(optimized) {
    let ui = SpreadsheetApp.getUi();

    let wordPrompt = ui.prompt(
        'Populate the field',
        'Word to correct',
        ui.ButtonSet.OK_CANCEL);

    const word = wordPrompt.getResponseText();

    if (word == null || word.length == 0) return;

    const gapCost = 1.0; //subject to change

    let candidates = getLocalDictionary();

    if (optimized) {
        candidates = candidates.filter((a) => Math.abs(a.length - word.length) <= 1) //is this proper or does it filter too much?
        candidates = candidates.filter((a) => a.charAt(0).toLowerCase() == word.charAt(0).toLowerCase());
    } else {
        candidates = candidates.filter((a) => Math.abs(a.length - word.length) <= 2) //is this proper or does it filter too much?
    }

    candidates.sort((a, b) => editDistance(word, a, gapCost)[1] - editDistance(word, b, gapCost)[1]);

    const topN = 20;
    let result = [];
    if (candidates != null && candidates.length >= topN) {
        result = candidates.slice(0, topN);
    }
    let title = "Autocorrect - " + word;
    var active = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = active.getSheetByName(title);
    if (sheet != null) {
        active.deleteSheet(sheet);
    }
    sheet = active.insertSheet();
    sheet.setName(title);
    sheet.appendRow(["Candidates", ...result]);
}

function correctWordTest(word, optimized) {
    const gapCost = 10.0; //subject to change
    console.log(word);
    let candidates = getLocalDictionary();
    if (optimized) {
        candidates = candidates.filter((a) => a.charAt(0).toLowerCase() == word.charAt(0).toLowerCase());
    }
    candidates.sort((a, b) => {
        return editDistance(word, a, gapCost)[1] - editDistance(word, b, gapCost)[1];
    });
    const topN = 3;
    let result = [];
    if (candidates != null && candidates.length >= topN) {
        result = candidates.slice(0, topN);
    }
    return result;
}

function setCharAt(string, char, index) {
    return string.substring(0, index) + char + string.substring(index + 1);
}

function test() {
    let result = editDistance("mean", "name", Number("2"));
    console.log(result[1]);
    let matrix = result[0];
    for (let i = matrix.length - 1; i >= 0; i--) {
        let currentRow = matrix[i];
        console.log(currentRow);
    }
    // console.log(setShortestPathParents(result[0], result[3]));
    // console.log(setCharAt("dog", "a", 1) == 'dag');
    // console.log(correctWordTest("ocurrance", true));
    // console.log(correctWordTest("evasipn", true))
}

function getLocalDictionary() {
  let emptyDict = []; //Dictionary omittied to keep file size small
  return emptyDict;
}
