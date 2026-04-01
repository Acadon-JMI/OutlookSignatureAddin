// signatureComposer.js - Block-based signature composition with layout marker support

var LAYOUT_START_TO_END = {
  'layout_logo_start':    'layout_logo_end',
  'layout_socials_start': 'layout_socials_end'
};

async function composeSignature(signatureObj, format, userData) {
  // signatureObj: { id, name, language, type, blocks: [{blockId}], customBlocks: [{id, name, htmlContent}] }
  // format: 'htm' or 'txt'
  // userData: merged user data with overrides applied

  var sections = [];
  var currentPlain = [];
  var inLayout = false;
  var currentEndId = null;
  var currentLogoHtml = '';
  var currentRightParts = [];

  for (var i = 0; i < signatureObj.blocks.length; i++) {
    var blockId = signatureObj.blocks[i].blockId;

    // Check if this is a layout start marker
    if (LAYOUT_START_TO_END[blockId]) {
      if (currentPlain.length > 0) {
        sections.push({ type: 'plain', parts: currentPlain });
        currentPlain = [];
      }
      var startHtml = await _getBlockContent(blockId, format, signatureObj);
      currentLogoHtml = startHtml ? applyPlaceholders(startHtml, userData) : '';
      currentRightParts = [];
      inLayout = true;
      currentEndId = LAYOUT_START_TO_END[blockId];
      continue;
    }

    // Check if this is the end marker for the current layout section
    if (inLayout && blockId === currentEndId) {
      sections.push({ type: 'twocol', logoHtml: currentLogoHtml, rightParts: currentRightParts });
      inLayout = false;
      currentEndId = null;
      continue;
    }

    // Load block content
    var blockContent = await _getBlockContent(blockId, format, signatureObj);
    if (!blockContent) continue;

    var processed = applyPlaceholders(blockContent, userData);

    if (inLayout) {
      currentRightParts.push(processed);
    } else {
      currentPlain.push(processed);
    }
  }

  // Push any remaining plain section
  if (currentPlain.length > 0) {
    sections.push({ type: 'plain', parts: currentPlain });
  }

  // Assemble final output
  if (format === 'htm') {
    return _assembleHtml(sections);
  } else {
    return _assembleText(sections);
  }
}

async function _getBlockContent(blockId, format, signatureObj) {
  // Custom blocks: read from signatureObj.customBlocks
  if (blockId.indexOf('custom_') === 0) {
    if (format === 'txt') return '';
    return getCustomBlockHtml(signatureObj, blockId);
  }
  // Server blocks: use blockLoader
  return await getBlockHtml(blockId, format);
}

function _assembleHtml(sections) {
  var parts = [];

  for (var i = 0; i < sections.length; i++) {
    var s = sections[i];
    if (s.type === 'plain') {
      if (s.parts.length > 0) {
        parts.push(s.parts.join('\n'));
      }
    } else {
      // Two-column layout section
      var twoColumn = '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; margin-left:10px;">\n';
      twoColumn += '  <tr>\n';
      twoColumn += '    <td valign="top" style="width:170px; padding:0 23px 0 0;">\n';
      twoColumn += '      ' + s.logoHtml + '\n';
      twoColumn += '    </td>\n';
      twoColumn += '    <td valign="top" style="padding:0;">\n';
      twoColumn += '      ' + s.rightParts.join('\n      ') + '\n';
      twoColumn += '    </td>\n';
      twoColumn += '  </tr>\n';
      twoColumn += '</table>';
      parts.push(twoColumn);
    }
  }

  return '<div style="font-family:\'Arial\',sans-serif;">\n' + parts.join('\n') + '\n</div>';
}

function _assembleText(sections) {
  var all = [];
  for (var i = 0; i < sections.length; i++) {
    var s = sections[i];
    if (s.type === 'plain') {
      all = all.concat(s.parts);
    } else {
      all = all.concat(s.rightParts);
    }
  }
  return all.join('\n\n');
}

async function injectSignature(event, isReply, triggerSource) {
  try {
    triggerSource = triggerSource || (isReply ? "Reply (Event)" : "NewMessage");
    console.log('acadon Signatur: Triggered by', triggerSource);

    // 1. Get current item to check subject for fallback
    var item = Office.context.mailbox.item;

    // Fallback: If not marked as reply but subject looks like one
    if (!isReply) {
      const subjectPromise = new Promise((resolve) => {
        item.subject.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value);
          } else {
            resolve("");
          }
        });
      });
      
      var subject = await subjectPromise;
      if (subject) {
        var lowerSubj = subject.toLowerCase();
        // Check for common reply/forward prefixes
        if (lowerSubj.indexOf('aw:') === 0 || lowerSubj.indexOf('re:') === 0 || lowerSubj.indexOf('fw:') === 0 || lowerSubj.indexOf('wg:') === 0) {
          console.log('acadon Signatur: Fallback detected reply/forward via subject:', subject);
          isReply = true;
          triggerSource += " (Fallback)";
        }
      }
    }

    // 2. Load user data and preferences
    var userData = await getUserData();
    var prefs = getPreferencesOrDefaults(userData.officeLocation);

    // 2.5. Check if auto-insert is enabled
    if (prefs.autoInsertEnabled === false) {
      console.log('acadon Signatur: Auto-insert is disabled in preferences.');
      event.completed();
      return;
    }

    // 3. Get the assigned signature
    var assignmentKey = isReply ? 'reply' : 'newMessage';
    var sigId = prefs.assignments[assignmentKey];
    var signature = getSignatureById(prefs, sigId);

    if (!signature) {
      console.warn('acadon Signatur: No signature found for assignment:', assignmentKey);
      event.completed();
      return;
    }

    console.log('acadon Signatur: Inserting signature:', signature.name, ' (Source:', triggerSource, ')');

    // 4. Merge user data with overrides and company info (use signature's language)
    var sigLang = signature.language || 'DE';
    var mergedData = mergeUserData(userData, prefs.overrides, sigLang);

    // 5. Compose the HTML signature from blocks
    var htmlSignature = await composeSignature(signature, 'htm', mergedData);

    // 6. Inject via setSignatureAsync
    item.body.setSignatureAsync(
      htmlSignature,
      { coercionType: Office.CoercionType.Html },
      function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error('acadon Signatur: setSignatureAsync failed:', asyncResult.error.message);
        }
        event.completed();
      }
    );
  } catch (err) {
    console.error('acadon Signatur: Injection error:', err.message);
    event.completed();
  }
}
