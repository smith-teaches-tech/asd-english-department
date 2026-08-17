/**
 * ONE-TIME MERGE SCRIPT — ASD English Library
 * ------------------------------------------------------------------
 * Folds 93 duplicate/triplicate book records into 81 surviving titles
 * and adds their per-room copies together. Room assignments are kept:
 * every copy stays in the room it was already counted in.
 *
 * HOW TO RUN
 *   1. Open the spreadsheet -> Extensions -> Apps Script
 *   2. Paste this whole file into a NEW script file (do not overwrite
 *      your existing web-app code)
 *   3. Select mergeDuplicates from the function dropdown -> Run
 *   4. Approve permissions if asked. Check the log when it finishes.
 *
 * SAFETY: before changing anything it copies the Books and Inventory
 * tabs to "<name> BACKUP <timestamp>". If anything looks wrong, delete
 * the live tabs and rename the backups back.
 *
 * NOTE: normalises course labels too ("Eng 10" -> "English 10",
 * "AP Lit" -> "AP Literature and Composition") so they match the
 * dropdown in the web app. Set NORMALISE_COURSES = false to skip that.
 */

var NORMALISE_COURSES = true;

// survivor book_id -> [title, author, course]
var KEEP = {"mm4ge739ju97": ["And the Mountains Echoed","Hosseini, Khaled",""],"mm4ge739mzzt": ["The Scarlet Letter","Hawthorne, Nathaniel","English 11"],"mm4ge739u5gy": ["Great Short Works of Leo Tolstoy","Tolstoy, Leo",""],"mm4ge739rvz0": ["Tess of the D'Urbervilles","Hardy, Thomas",""],"mm4ge7jqpwoq": ["Best-Loved Folktales","Cole, Joanna",""],"mm4ge7393un5": ["Collected Stories","Márquez, Gabriel García",""],"mm4ge739pd5g": ["Uncle Tom's Cabin","Stowe, Harriet Beecher",""],"mm4ge8sw8yc6": ["Gothic Short Stories","Blair, David (ed.)",""],"mm4ge84sxgvw": ["One Hundred Years of Solitude","Gabriel Garcia Marquez",""],"mm4ge8as92g1": ["Salt, Sugar, Fat","Moss, Michael",""],"mm4ge739myj3": ["Poe: Selected Tales","Poe, Edgar Allan",""],"mm4ge8asdsam": ["The Handmaid's Tale","Atwood, Margaret","English 10"],"mm4ge73aglyd": ["The Great Gatsby","Fitzgerald, F. Scott","English 11"],"mm4ge84sl0g5": ["The Canterbury Tales","Geoffrey Chaucer",""],"mm4ge8asvcd7": ["50 Essays - Second Edition","Cohen, Samuel",""],"mm4ge7jq563y": ["Nectar in a Sieve","Markandaya, Kamala","English 12"],"mm4ge7jq9ptr": ["Cry, the Beloved Country","Paton, Alan","English 12"],"mm4ge7jqi3r0": ["Death of a Salesman","Arthur Miller","English 11"],"mm4ge8gir64s": ["Their Eyes Were Watching God","Hurston, Zora Neale","English 11"],"mm4ge8as3ipi": ["In Cold Blood","Capote, Truman","AP Language and Composition"],"mm4ge8atiwhv": ["The Bluest Eye","Morrison, Toni",""],"mm4ge7pzkbl7": ["Lord of the Flies","William Golding","English 9"],"mm4ge7pzsgs0": ["To Kill a Mockingbird","Lee, Harper","English 9"],"mm4ge7pzpic2": ["Of Mice and Men","Steinbeck, John","English 9"],"mm4ge7pz4jzb": ["Fahrenheit 451","Bradbury, Ray","English 10"],"mm4ge7pz5ybz": ["Invisible Man","Ellison, Ralph","AP Literature and Composition"],"mm4ge7pzxkob": ["Macbeth","Shakespeare, William","English 10"],"mm4ge8asudlr": ["The Language of Composition","Renee H. Shea & Lawrence Scanlon & Robin Dissin Aufses",""],"mm4ge84s5efp": ["The Road","McCarthy, Cormac","English 10"],"mm4ge84sm1dm": ["In the Time of the Butterflies","Julia Alvarez","English 10"],"mm4ge8swkhoi": ["Things Fall Apart","Chinua Achebe","AP Literature and Composition"],"mm4ge84sup8b": ["The Picture of Dorian Gray","Wilde, Oscar",""],"mm4ge84sxgqr": ["Black Boy","Wright, Richard",""],"mm4ge84s6kow": ["Jane Eyre","Charlotte Bronte",""],"mm4ge8swtf84": ["The Complete Stories of Flannery O'Connor","O'Connor, Flannery",""],"mm4ge84s5v8w": ["Hamlet","Shakespeare, William","AP Literature and Composition"],"mm4ge84t62sy": ["Romeo and Juliet","Shakespeare, William",""],"mm4ge8swownz": ["Pride and Prejudice","Jane Austen",""],"mm4ge8assse8": ["1984","Orwell, George","English 10"],"mm4ge8asmdhs": ["Animal Farm","Orwell, George",""],"mm4ge8asqe8x": ["I Know Why the Caged Bird Sings","Angelou, Maya",""],"mm4ge8asxre0": ["The Mother Tongue","Bryson, Bill",""],"mm4ge8gip57s": ["Outliers","Gladwell, Malcolm","AP Language and Composition"],"mm4ge8as8rqn": ["The Shallows: How the Internet Is Changing Us","Carr, Nicholas",""],"mm4ge8ascwwq": ["The Beauty Myth","Wolf, Naomi",""],"mm4ge8asd0hh": ["Black Swan Green","Mitchell, David",""],"mm4ge8as0134": ["Influence: The Power of Persuasion","Cialdini, Robert",""],"mm4ge8gir0i5": ["Uglies","Westerfeld, Scott",""],"mm4ge8asj8ka": ["Quiet: The Power of Introverts","Cain, Susan",""],"mm4ge8as4724": ["Stiff","Roach, Mary",""],"mm4ge8as5qq4": ["Amusing Ourselves to Death","Postman, Neil",""],"mm4ge8mlt688": ["The Walmart Effect","Fishman, Charles",""],"mm4ge8mlpw8k": ["Nickel and Dimed","Ehrenreich, Barbara",""],"mm4ge8asgfma": ["Wonder","Palacio, R.J.",""],"mm4ge8atrzxe": ["Born a Crime","Noah, Trevor",""],"mm4ge8swzy29": ["All the Light We Cannot See","Doerr, Anthony",""],"mm4ge8gii12z": ["Children of Blood and Bone","Adeyemi, Tomi",""],"mm4ge8ml3ay8": ["Buy-Ology","Lindstrom, Martin",""],"mm4ge8atqh5v": ["Hillbilly Elegy","Vance, JD",""],"mm4ge8mlrj4t": ["The Wind-Up Bird Chronicle","Murakami, Haruki",""],"mm4ge8gi0k0g": ["The Last Lecture","Pausch, Randy",""],"mm4ge8gi0fkk": ["The Joy Luck Club","Tan, Amy",""],"mm4ge8ml5ans": ["The Illustrated Man","Bradbury, Ray",""],"mm4ge8giocdh": ["The No. 1 Ladies' Detective Agency","Smith, Alexander McCall",""],"mm4ge8swtce1": ["The Secret Side of Empty","Andreu, Maria E.",""],"mm4ge8swy4jg": ["The Diary of Anne Frank","Frank, Anne",""],"mm4ge8swdjke": ["Into the Wild","Krakauer, Jack",""],"mm4ge8gipf4s": ["This Land Is Our Land","Mehta, Suketu",""],"mm4ge8mliot8": ["The Etymologicon","Forsyth, Mark",""],"mm4ge8giozr7": ["Dear Martin","Stone, Nic",""],"mm4ge8gia2iz": ["Salt to the Sea","Sepetys, Ruta",""],"mm4ge8giky0j": ["All American Boys","Brendan Kiely and Jason Reynolds",""],"mm4ge8swtwmv": ["The Secret Life of Bees","Kidd, Sue Monk",""],"mm4ge8ml008j": ["The Armchair Economist","Landsburg, Steven E.",""],"mm4ge8gi6hhf": ["The Hate U Give","Thomas, Angie",""],"mm4ge8sw2ykl": ["A Step From Heaven","Na, An",""],"mm4ge8sw16y1": ["A Mother's Reckoning","Klebold, Sue",""],"mm4ge8mlufy1": ["Frankenstein","Shelley, Mary","AP Literature and Composition"],"mm4ge8mlmxez": ["The Crucible","Miller, Arthur","English 11"],"mm4ge8swhoop": ["Brave New World","Huxley, Aldous",""],"mm4ge8ml6ult": ["Nimona","Noelle Stevenson",""]};

// duplicate book_id -> survivor book_id it folds into
var MERGE_MAP = {"mm4ge8swxagw": "mm4ge739ju97","mm4ge8atdkyo": "mm4ge739mzzt","mm4ge8mlpbp0": "mm4ge739u5gy","mm4ge8sw8lcu": "mm4ge739u5gy","mm4ge8at3krb": "mm4ge739u5gy","mm4ge8ml7ftn": "mm4ge739rvz0","mm4ge8atgtjn": "mm4ge739rvz0","mm4ge8swuqv9": "mm4ge739rvz0","mm4ge73969cv": "mm4ge7jqpwoq","mm4ge8at72tb": "mm4ge7393un5","mm4ge8atqrdf": "mm4ge739pd5g","mm4ge739u8vr": "mm4ge8sw8yc6","mm4ge739nrss": "mm4ge84sxgvw","mm4ge739v1um": "mm4ge8as92g1","mm4ge8mllffn": "mm4ge8as92g1","mm4ge8as4o7a": "mm4ge739myj3","mm4ge73alx32": "mm4ge8asdsam","mm4ge8atxcis": "mm4ge73aglyd","mm4ge73akgno": "mm4ge84sl0g5","mm4ge73algs3": "mm4ge8asvcd7","mm4ge8atqe18": "mm4ge7jq563y","mm4ge7jqn6e6": "mm4ge7jq9ptr","mm4ge8swaugw": "mm4ge7jqi3r0","mm4ge7jqc1m7": "mm4ge8gir64s","mm4ge8auaeo6": "mm4ge8gir64s","mm4ge7jqb2fe": "mm4ge8as3ipi","mm4ge8swc5il": "mm4ge8as3ipi","mm4ge7jq0bva": "mm4ge8atiwhv","mm4ge8gixe4s": "mm4ge7pzkbl7","mm4ge84shvtl": "mm4ge7pzsgs0","mm4ge8atkeaj": "mm4ge7pzsgs0","mm4ge8auc49n": "mm4ge7pzpic2","mm4ge84svxjm": "mm4ge7pz4jzb","mm4ge8asdiac": "mm4ge7pz4jzb","mm4ge8sw64le": "mm4ge7pz4jzb","mm4ge8atovyk": "mm4ge7pz5ybz","mm4ge8swpv81": "mm4ge7pz5ybz","mm4ge8asg2k3": "mm4ge7pzxkob","mm4ge7pzx4z2": "mm4ge8asudlr","mm4ge8mlftbr": "mm4ge84s5efp","mm4ge8mlqlq9": "mm4ge84sm1dm","mm4ge84s73ns": "mm4ge8swkhoi","mm4ge8aunot7": "mm4ge84sup8b","mm4ge8at5icc": "mm4ge84sxgqr","mm4ge84sql9w": "mm4ge84s6kow","mm4ge84s6taz": "mm4ge8swtf84","mm4ge8asut4c": "mm4ge84s5v8w","mm4ge8au72wp": "mm4ge84t62sy","mm4ge84t51pr": "mm4ge8swownz","mm4ge8gi2pes": "mm4ge8assse8","mm4ge8gixanv": "mm4ge8asmdhs","mm4ge8swz3za": "mm4ge8asqe8x","mm4ge8mlxur6": "mm4ge8asxre0","mm4ge8asebiv": "mm4ge8gip57s","mm4ge8mlaknf": "mm4ge8as8rqn","mm4ge8mldrlr": "mm4ge8ascwwq","mm4ge8swmt0g": "mm4ge8asd0hh","mm4ge8mlw696": "mm4ge8as0134","mm4ge8asaub9": "mm4ge8gir0i5","mm4ge8gizofs": "mm4ge8asj8ka","mm4ge8mlkcmw": "mm4ge8as4724","mm4ge8mlv6t2": "mm4ge8as5qq4","mm4ge8ase8k9": "mm4ge8mlt688","mm4ge8asqg0z": "mm4ge8mlpw8k","mm4ge8gia9me": "mm4ge8asgfma","mm4ge8mlvmip": "mm4ge8atrzxe","mm4ge8ataur3": "mm4ge8swzy29","mm4ge8atisoe": "mm4ge8gii12z","mm4ge8atzdav": "mm4ge8ml3ay8","mm4ge8ml1gz9": "mm4ge8atqh5v","mm4ge8atiizo": "mm4ge8mlrj4t","mm4ge8atgsyz": "mm4ge8gi0k0g","mm4ge8at9bj9": "mm4ge8gi0fkk","mm4ge8atr44m": "mm4ge8ml5ans","mm4ge8atskt1": "mm4ge8giocdh","mm4ge8at4neh": "mm4ge8swtce1","mm4ge8atolzc": "mm4ge8swy4jg","mm4ge8atgqxg": "mm4ge8swdjke","mm4ge8atyzdg": "mm4ge8gipf4s","mm4ge8sw71yt": "mm4ge8gipf4s","mm4ge8at6cnp": "mm4ge8mliot8","mm4ge8at5zeu": "mm4ge8giozr7","mm4ge8ata821": "mm4ge8gia2iz","mm4ge8atm434": "mm4ge8giky0j","mm4ge8atzemm": "mm4ge8swtwmv","mm4ge8atmkbo": "mm4ge8ml008j","mm4ge8atxf7l": "mm4ge8gi6hhf","mm4ge8at8lcz": "mm4ge8sw2ykl","mm4ge8atn3b8": "mm4ge8sw16y1","mm4ge8at6k1p": "mm4ge8mlufy1","mm4ge8auag6u": "mm4ge8mlmxez","mm4ge8au0n81": "mm4ge8swhoop","mm4ge8gizj3w": "mm4ge8ml6ult"};

function mergeDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var booksSheet = findSheet_(ss, ['book_id', 'title', 'author']);
  var invSheet   = findSheet_(ss, ['room_id', 'book_id', 'quantity']);
  if (!booksSheet) throw new Error('Could not find the Books tab (needs book_id/title/author headers).');
  if (!invSheet)   throw new Error('Could not find the Inventory tab (needs room_id/book_id/quantity headers).');

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM-dd HH.mm');
  booksSheet.copyTo(ss).setName(booksSheet.getName() + ' BACKUP ' + stamp);
  invSheet.copyTo(ss).setName(invSheet.getName() + ' BACKUP ' + stamp);
  Logger.log('Backups created.');

  // ---------- INVENTORY: add duplicate copies into the survivor ----------
  var inv = invSheet.getDataRange().getValues();
  var iHdr = inv[0].map(norm_);
  var iRoom = iHdr.indexOf('room_id'), iBook = iHdr.indexOf('book_id');
  var iQty = iHdr.indexOf('quantity'), iUpd = iHdr.indexOf('updated_at');

  var totalBefore = 0;
  var bucket = {};   // "room|survivorId" -> {room, book, qty, upd}
  var order  = [];
  for (var r = 1; r < inv.length; r++) {
    var room = String(inv[r][iRoom] || '').trim();
    var book = String(inv[r][iBook] || '').trim();
    if (!room || !book) continue;
    var qty = Number(inv[r][iQty]) || 0;
    totalBefore += qty;
    var target = MERGE_MAP[book] || book;          // fold duplicates in
    var key = room + '|' + target;
    if (!bucket[key]) { bucket[key] = {room: room, book: target, qty: 0, upd: ''}; order.push(key); }
    bucket[key].qty += qty;                        // <- the addition
    var upd = inv[r][iUpd] ? String(inv[r][iUpd]) : '';
    if (upd > bucket[key].upd) bucket[key].upd = upd;
  }

  var outRows = [], totalAfter = 0;
  for (var i = 0; i < order.length; i++) {
    var b = bucket[order[i]];
    if (b.qty <= 0) continue;                      // drop empty rows
    totalAfter += b.qty;
    var row = [];
    row[iRoom] = b.room; row[iBook] = b.book; row[iQty] = b.qty;
    if (iUpd >= 0) row[iUpd] = b.upd || new Date().toISOString();
    for (var c = 0; c < iHdr.length; c++) if (row[c] === undefined) row[c] = '';
    outRows.push(row);
  }
  if (totalBefore !== totalAfter) {
    throw new Error('ABORTED - copy count changed (' + totalBefore + ' -> ' + totalAfter + '). Nothing was saved.');
  }

  if (invSheet.getLastRow() > 1) {
    invSheet.getRange(2, 1, invSheet.getLastRow() - 1, iHdr.length).clearContent();
  }
  if (outRows.length) invSheet.getRange(2, 1, outRows.length, iHdr.length).setValues(outRows);
  Logger.log('Inventory: ' + (inv.length - 1) + ' rows -> ' + outRows.length + ' rows. ' + totalAfter + ' copies (unchanged).');

  // ---------- BOOKS: retitle survivors, delete duplicates ----------
  var bk = booksSheet.getDataRange().getValues();
  var bHdr = bk[0].map(norm_);
  var bId = bHdr.indexOf('book_id'), bTi = bHdr.indexOf('title');
  var bAu = bHdr.indexOf('author'),  bCo = bHdr.indexOf('course');
  var bUp = bHdr.indexOf('updated_at');

  var kept = [], removed = 0, retitled = 0, recoursed = 0;
  for (var r = 1; r < bk.length; r++) {
    var id = String(bk[r][bId] || '').trim();
    if (!id) continue;
    if (MERGE_MAP[id]) { removed++; continue; }            // duplicate - drop
    var row = bk[r].slice();
    if (KEEP[id]) {                                        // survivor - clean up
      row[bTi] = KEEP[id][0];
      row[bAu] = KEEP[id][1];
      if (bCo >= 0) row[bCo] = KEEP[id][2];
      if (bUp >= 0) row[bUp] = new Date().toISOString();
      retitled++;
    } else if (NORMALISE_COURSES && bCo >= 0) {
      var fixed = fixCourse_(row[bCo]);
      if (fixed !== row[bCo]) { row[bCo] = fixed; recoursed++; }
    }
    kept.push(row);
  }

  if (booksSheet.getLastRow() > 1) {
    booksSheet.getRange(2, 1, booksSheet.getLastRow() - 1, bHdr.length).clearContent();
  }
  if (kept.length) booksSheet.getRange(2, 1, kept.length, bHdr.length).setValues(kept);

  var msg = 'DONE. ' + removed + ' duplicate records folded into ' + retitled +
            ' titles. Catalogue: ' + (bk.length - 1) + ' -> ' + kept.length + ' books. ' +
            totalAfter + ' copies preserved. ' + recoursed + ' course labels tidied.';
  Logger.log(msg);
  SpreadsheetApp.getActive().toast(msg, 'Merge complete', 30);
}

function norm_(h) { return String(h || '').trim().toLowerCase(); }

function findSheet_(ss, needed) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (/BACKUP/i.test(sheets[i].getName())) continue;
    if (sheets[i].getLastRow() < 1) continue;
    var hdr = sheets[i].getRange(1, 1, 1, Math.max(1, sheets[i].getLastColumn())).getValues()[0].map(norm_);
    var ok = true;
    for (var j = 0; j < needed.length; j++) if (hdr.indexOf(needed[j]) < 0) ok = false;
    if (ok) return sheets[i];
  }
  return null;
}

function fixCourse_(c) {
  var v = String(c || '').trim();
  var map = {
    'eng 9': 'English 9', 'eng 10': 'English 10', 'eng 11': 'English 11', 'eng 12': 'English 12',
    'english9': 'English 9', 'ap lang': 'AP Language and Composition',
    'ap lit': 'AP Literature and Composition', 'unassigned': ''
  };
  var hit = map[v.toLowerCase()];
  return hit === undefined ? v : hit;
}
