if (D2RMM.getVersion == null || D2RMM.getVersion() < 1.6) {
  D2RMM.error('Requires D2RMM version 1.6 or higher.');
  return;
}

function getTabName(value, defaultValue) {
  return !config.isCustomTabsEnabled || value === '' ? defaultValue : value;
}

const STASH_TAB_NAMES = [
  getTabName(config.tabNamePersonal, '@personal'),
  getTabName(config.tabNameShared1, '@shared'),
  getTabName(config.tabNameShared2, '@shared'),
  getTabName(config.tabNameShared3, '@shared'),
  getTabName(config.tabNameShared4, '@shared'),
  getTabName(config.tabNameShared5, '@shared'),
  getTabName(config.tabNameShared6, '@shared'),
  getTabName(config.tabNameShared7, '@shared'),
];

const inventoryFilename = 'global\\excel\\inventory.txt';
const inventory = D2RMM.readTsv(inventoryFilename);
inventory.rows.forEach((row) => {
  const id = row.class;
  if (
    id === 'Bank Page 1' ||
    id === 'Big Bank Page 1' ||
    id === 'Bank Page2' ||
    id === 'Big Bank Page2'
  ) {
    row.gridX = 16;
    row.gridY = 13;
  }
});
D2RMM.writeTsv(inventoryFilename, inventory);

const profileHDFilename = 'global\\ui\\layouts\\_profilehd.json';
const profileHD = D2RMM.readJson(profileHDFilename);
profileHD.LeftPanelRect_ExpandedStash = {
  x: 236,
  y: -651,
  width: 1687,
  height: 1507,
};
profileHD.PanelClickCatcherRect_ExpandedStash = {
  x: 0,
  y: 0,
  width: 1687,
  height: 1507,
};
// offset the left hinge so that it doesn't overlap with content of the left panel
profileHD.LeftHingeRect = { x: -236 - 25, y: 630 };
D2RMM.writeJson(profileHDFilename, profileHD);

const profileLVFilename = 'global\\ui\\layouts\\_profilelv.json';
const profileLV = D2RMM.readJson(profileLVFilename);
profileLV.LeftPanelRect_ExpandedStash = {
  x: 0,
  y: -856,
  width: 1687,
  height: 1507,
  scale: 1.16,
};
D2RMM.writeJson(profileLVFilename, profileLV);

const bankOriginalLayoutFilename =
  'global\\ui\\layouts\\bankoriginallayout.json';
const bankOriginalLayout = D2RMM.readJson(bankOriginalLayoutFilename);
// TODO: new sprite & layout for classic UI
bankOriginalLayout.children.forEach((child) => {
  if (child.name === 'grid') {
    child.fields.cellCount.x = 16;
    child.fields.cellCount.y = 13;
  }
});
D2RMM.writeJson(bankOriginalLayoutFilename, bankOriginalLayout);

const bankExpansionLayoutFilename =
  'global\\ui\\layouts\\bankexpansionlayout.json';
const bankExpansionLayout = D2RMM.readJson(bankExpansionLayoutFilename);
// TODO: new sprite & layout for classic UI
bankExpansionLayout.children = bankExpansionLayout.children.map((child) => {
  if (
    child.name === 'PreviousSeasonToggleDisplay' ||
    child.name === 'PreviousLadderSeasonBankTabs'
  ) {
    return false;
  }
  if (child.name === 'grid') {
    child.fields.cellCount.x = 16;
    child.fields.cellCount.y = 13;
  }
  if (child.name === 'BankTabs') {
    child.fields.tabCount = 8;
    child.fields.textStrings = STASH_TAB_NAMES;
  }
  return true;
});
D2RMM.writeJson(bankExpansionLayoutFilename, bankExpansionLayout);

const bankOriginalLayoutHDFilename =
  'global\\ui\\layouts\\bankoriginallayouthd.json';
const bankOriginalLayoutHD = D2RMM.readJson(bankOriginalLayoutHDFilename);
bankOriginalLayoutHD.fields.rect = '$LeftPanelRect_ExpandedStash';
bankOriginalLayoutHD.children.forEach((child) => {
  if (child.name === 'grid') {
    child.fields.cellCount.x = 16;
    child.fields.cellCount.y = 13;
    child.fields.rect.x = child.fields.rect.x - 229;
    child.fields.rect.y = child.fields.rect.y - 572;
  }
  if (child.name === 'click_catcher') {
    child.fields.rect = '$PanelClickCatcherRect_ExpandedStash';
  }
  if (child.name === 'background') {
    child.fields.filename = 'PANEL\\Stash\\StashPanel_BG_Expanded';
    child.fields.rect = { x: 0, y: 0 };
  }
  if (child.name === 'gold_amount') {
    child.fields.rect.x = 60 + 60 + 16;
    child.fields.rect.y = 61 + 16;
  }
  if (child.name === 'gold_withdraw') {
    child.fields.rect.x = 60 + 16;
    child.fields.rect.y = 58 + 16;
  }
  if (child.name === 'title') {
    child.fields.rect = {
      x: 91 + (1687 - 1162) / 2,
      y: 64,
      width: 972,
      height: 71,
    };
  }
  if (child.name === 'close') {
    child.fields.rect.x = child.fields.rect.x + (1687 - 1162);
  }
});
D2RMM.writeJson(bankOriginalLayoutHDFilename, bankOriginalLayoutHD);

const bankExpansionLayoutHDFilename =
  'global\\ui\\layouts\\bankexpansionlayouthd.json';
const bankExpansionLayoutHD = D2RMM.readJson(bankExpansionLayoutHDFilename);
bankExpansionLayoutHD.children = bankExpansionLayoutHD.children.filter(
  (child) => {
    if (
      child.name === 'PreviousSeasonToggleDisplay' ||
      child.name === 'PreviousLadderSeasonBankTabs'
    ) {
      return false;
    }
    if (child.name === 'grid') {
      child.fields.cellCount.x = 16;
      child.fields.cellCount.y = 13;
      child.fields.rect.x = child.fields.rect.x - 37;
      child.fields.rect.y = child.fields.rect.y - 58;
    }
    if (child.name === 'background') {
      child.fields.filename = 'PANEL\\Stash\\StashPanel_BG_Expanded';
    }
    if (child.name === 'BankTabs') {
      child.fields.filename = 'PANEL\\stash\\Stash_Tabs_Expanded';
      child.fields.rect.x = child.fields.rect.x - 30;
      child.fields.rect.y = child.fields.rect.y - 56;
      child.fields.tabCount = 8;
      // 249 x 80 -> 197 x 80 (bottom 5 pixels are overlay)
      child.fields.tabSize = { x: 197, y: 75 };
      child.fields.tabPadding = { x: 0, y: 0 };
      child.fields.inactiveFrames = [0, 0, 0, 0, 0, 0, 0, 0];
      child.fields.activeFrames = [1, 1, 1, 1, 1, 1, 1, 1];
      child.fields.disabledFrames = [0, 0, 0, 0, 0, 0, 0, 0];
      child.fields.textStrings = STASH_TAB_NAMES;
    }
    if (child.name === 'gold_amount') {
      child.fields.rect.x = 60 + 60;
      child.fields.rect.y = 61;
    }
    if (child.name === 'gold_withdraw') {
      child.fields.rect.x = 60;
      child.fields.rect.y = 58;
    }
    if (child.name === 'title') {
      // hide title
      return false;
    }
    return true;
  }
);
D2RMM.writeJson(bankExpansionLayoutHDFilename, bankExpansionLayoutHD);

const controllerOverlayHDFilename =
  'global\\ui\\layouts\\controller\\controlleroverlayhd.json';
const controllerOverlayHD = D2RMM.readJson(controllerOverlayHDFilename);
controllerOverlayHD.children.forEach((child) => {
  if (child.name === 'Anchor') {
    child.children.forEach((subchild) => {
      if (subchild.name === 'ControllerCursorBounds') {
        delete subchild.fields.fitToParent;
        subchild.fields.rect = {
          x: -285,
          y: 0,
          width: 2880 + 285,
          height: 1763,
        };
      }
    });
  }
});
D2RMM.writeJson(controllerOverlayHDFilename, controllerOverlayHD);

const bankOriginalControllerLayoutHDFilename =
  'global\\ui\\layouts\\controller\\bankoriginallayouthd.json';
const bankOriginalControllerLayoutHD = D2RMM.readJson(
  bankOriginalControllerLayoutHDFilename
);
bankOriginalControllerLayoutHD.children.forEach((child) => {
  if (child.name === 'background') {
    child.fields.filename =
      'Controller/Panel/Stash/V2/Classic_StashPanelBG_Expanded';
    child.fields.rect.x = child.fields.rect.x - 285 - 81 - 2 - 120;
    child.fields.rect.y = child.fields.rect.y + 17 - 293 + 100;
  }
  if (child.name === 'gold_amount' || child.name === 'gold_withdraw') {
    child.fields.rect.x = child.fields.rect.x - 476 - 280;
    child.fields.rect.y = child.fields.rect.y - 1404 + 30;
  }
  if (child.name === 'gold_max') {
    child.fields.rect.x = child.fields.rect.x - 476 + 927;
    child.fields.rect.y = child.fields.rect.y - 1404 - 90 + 25;
  }
  if (child.name === 'grid') {
    child.fields.rect.x = -285 + 9;
    child.fields.rect.y = 119;
  }
});
D2RMM.writeJson(
  bankOriginalControllerLayoutHDFilename,
  bankOriginalControllerLayoutHD
);

const bankExpansionControllerLayoutHDFilename =
  'global\\ui\\layouts\\controller\\bankexpansionlayouthd.json';
const bankExpansionControllerLayoutHD = D2RMM.readJson(
  bankExpansionControllerLayoutHDFilename
);
bankExpansionControllerLayoutHD.children =
  bankExpansionControllerLayoutHD.children.filter((child) => {
    if (
      child.name === 'PreviousSeasonToggleDisplay' ||
      child.name === 'PreviousLadderSeasonBankTabs'
    ) {
      return false;
    }
    if (child.name === 'background') {
      child.fields.filename = 'Controller/Panel/Stash/V2/StashPanelBG_Expanded';
      child.fields.rect.x = child.fields.rect.x - 285 - 81;
      child.fields.rect.y = child.fields.rect.y + 17 - 293;
    }
    if (child.name === 'gold_amount' || child.name === 'gold_withdraw') {
      child.fields.rect.x = child.fields.rect.x - 476 - 280;
      child.fields.rect.y = child.fields.rect.y - 1404;
    }
    if (child.name === 'gold_max') {
      child.fields.rect.x = child.fields.rect.x - 476 + 927;
      child.fields.rect.y = child.fields.rect.y - 1404 - 90;
    }
    if (child.name === 'grid') {
      child.fields.cellCount.x = 16;
      child.fields.cellCount.y = 13;
      child.fields.rect.x = -285 + 9;
      child.fields.rect.y = 119;
    }
    if (child.name === 'BankTabs') {
      child.fields.filename = 'Controller/Panel/Stash/V2/StashTabs_Expanded';
      child.fields.focusIndicatorFilename =
        'Controller/HoverImages/StashTab_Hover_Expanded';
      child.fields.rect.x = child.fields.rect.x - 300;
      child.fields.rect.y = child.fields.rect.y + 10;
      child.fields.tabCount = 8;
      child.fields.tabSize = { x: 175, y: 120 };
      child.fields.tabPadding = { x: 0, y: 0 };
      child.fields.inactiveFrames = [1, 1, 1, 1, 1, 1, 1, 1];
      child.fields.activeFrames = [0, 0, 0, 0, 0, 0, 0, 0];
      child.fields.disabledFrames = [1, 1, 1, 1, 1, 1, 1, 1];
      child.fields.textStrings = STASH_TAB_NAMES;
      child.fields.tabLeftIndicatorPosition = { x: -42, y: -2 };
      child.fields.tabRightIndicatorPosition = { x: 1135 + 300, y: -2 };
    }
    return true;
  });
D2RMM.writeJson(
  bankExpansionControllerLayoutHDFilename,
  bankExpansionControllerLayoutHD
);

{
  const inventoryFilename = 'global\\excel\\inventory.txt';
  const inventory = D2RMM.readTsv(inventoryFilename);
  inventory.rows.forEach((row) => {
    const id = row.class;
    const classes = [
      'Amazon',
      'Assassin',
      'Barbarian',
      'Druid',
      'Necromancer',
      'Paladin',
      'Sorceress',
    ];
    if (
      classes.indexOf(id) !== -1 ||
      classes.map((cls) => `${cls}2`).indexOf(id) !== -1
    ) {
      row.gridX = 13;
      row.gridY = 8;
    }
  });
  D2RMM.writeTsv(inventoryFilename, inventory);

  const profileHDFilename = 'global\\ui\\layouts\\_profilehd.json';
  const profileHD = D2RMM.readJson(profileHDFilename);
  profileHD.RightPanelRect_ExpandedInventory = {
    x: -1394 - (1382 - 1162),
    y: -651,
    width: 1382,
    height: 1507,
  };
  profileHD.PanelClickCatcherRect_ExpandedInventory = {
    x: 0,
    y: 0,
    width: 1162,
    height: 1507,
  };
  // offset the right hinge so that it doesn't overlap with content of the right panel
  profileHD.RightHingeRect = { x: 1076 + 20, y: 630 };
  profileHD.RightHingeRect_ExpandedInventory = {
    x: 1076 + (1382 - 1162) + 20,
    y: 630,
  };
  D2RMM.writeJson(profileHDFilename, profileHD);

  const profileLVFilename = 'global\\ui\\layouts\\_profilelv.json';
  const profileLV = D2RMM.readJson(profileLVFilename);
  profileLV.RightPanelRect_ExpandedInventory = {
    x: -1346 - (1382 - 1162) * 1.16,
    y: -856,
    width: 1382,
    height: 1507,
    scale: 1.16,
  };
  D2RMM.writeJson(profileLVFilename, profileLV);

  const playerInventoryOriginalLayoutFilename =
    'global\\ui\\layouts\\playerinventoryoriginallayout.json';
  const playerInventoryOriginalLayout = D2RMM.readJson(
    playerInventoryOriginalLayoutFilename
  );
  // TODO: new sprite & layout for classic UI
  playerInventoryOriginalLayout.children.forEach((child) => {
    if (child.name === 'grid') {
      child.fields.cellCount.x = 13;
      child.fields.cellCount.y = 8;
    }
  });
  D2RMM.writeJson(
    playerInventoryOriginalLayoutFilename,
    playerInventoryOriginalLayout
  );

  const playerInventoryOriginalLayoutHDFilename =
    'global\\ui\\layouts\\playerinventoryoriginallayouthd.json';
  const playerInventoryOriginalLayoutHD = D2RMM.readJson(
    playerInventoryOriginalLayoutHDFilename
  );
  playerInventoryOriginalLayoutHD.fields.rect =
    '$RightPanelRect_ExpandedInventory';
  playerInventoryOriginalLayoutHD.children =
    playerInventoryOriginalLayoutHD.children.filter((child) => {
      if (child.name === 'background') {
        child.fields.filename = 'PANEL\\Inventory\\Classic_Background_Expanded';
      }
      if (child.name === 'click_catcher') {
        child.fields.rect = { x: 0, y: 45, width: 1093, height: 1495 };
      }
      if (child.name === 'RightHinge') {
        child.fields.rect = '$RightHingeRect_ExpandedInventory';
      }
      if (child.name === 'title') {
        child.fields.rect = {
          x: 91 + (1382 - 1162) / 2,
          y: 64,
          width: 972,
          height: 71,
        };
      }
      if (child.name === 'close') {
        child.fields.rect.x = child.fields.rect.x + (1382 - 1162);
      }
      if (child.name === 'grid') {
        child.fields.cellCount.x = 13;
        child.fields.cellCount.y = 8;
        child.fields.rect.x = child.fields.rect.x - 37;
        child.fields.rect.y = child.fields.rect.y - 229;
      }
      if (child.name === 'slot_right_arm') {
        child.fields.rect.x = child.fields.rect.x - 14;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      if (child.name === 'slot_left_arm') {
        child.fields.rect.x = child.fields.rect.x + 227;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      if (child.name === 'slot_torso') {
        child.fields.rect.x = child.fields.rect.x + 101;
        child.fields.rect.y = child.fields.rect.y - 229;
      }
      if (child.name === 'slot_head') {
        child.fields.rect.x = child.fields.rect.x - 144;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      if (child.name === 'slot_gloves') {
        child.fields.rect.x = child.fields.rect.x + 231;
        child.fields.rect.y = child.fields.rect.y - 233;
      }
      if (child.name === 'slot_feet') {
        child.fields.rect.x = child.fields.rect.x - 26;
        child.fields.rect.y = child.fields.rect.y - 231;
      }
      if (child.name === 'slot_belt') {
        child.fields.rect.x = child.fields.rect.x + 101;
        child.fields.rect.y = child.fields.rect.y - 234;
      }
      if (child.name === 'slot_neck') {
        child.fields.rect.x = child.fields.rect.x + 99;
        child.fields.rect.y = child.fields.rect.y - 182;
      }
      if (child.name === 'slot_right_hand') {
        child.fields.rect.x = child.fields.rect.x + 474;
        child.fields.rect.y = child.fields.rect.y - 466;
      }
      if (child.name === 'slot_left_hand') {
        child.fields.rect.x = child.fields.rect.x + 232;
        child.fields.rect.y = child.fields.rect.y - 466;
      }
      if (child.name === 'gold_amount' || child.name === 'gold_button') {
        child.fields.rect.x = child.fields.rect.x - 291;
        child.fields.rect.y = child.fields.rect.y - 1267;
      }
      return true;
    });
  D2RMM.writeJson(
    playerInventoryOriginalLayoutHDFilename,
    playerInventoryOriginalLayoutHD
  );

  const playerInventoryExpansionLayoutHDFilename =
    'global\\ui\\layouts\\playerinventoryexpansionlayouthd.json';
  const playerInventoryExpansionLayoutHD = D2RMM.readJson(
    playerInventoryExpansionLayoutHDFilename
  );
  playerInventoryExpansionLayoutHD.children =
    playerInventoryExpansionLayoutHD.children.filter((child) => {
      if (child.name === 'click_catcher') {
        // make click catcher work the same way as in the originallayouthd file
        return false;
      }
      if (child.name === 'background') {
        child.fields.filename = 'PANEL\\Inventory\\Background_Expanded';
      }
      if (
        child.name === 'background_right_arm' ||
        child.name === 'background_right_arm_selected' ||
        child.name === 'weaponswap_right_arm' ||
        child.name === 'text_i_left' ||
        child.name === 'text_ii_left'
      ) {
        child.fields.rect.x = child.fields.rect.x - 14;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      if (
        child.name === 'background_left_arm' ||
        child.name === 'background_left_arm_selected' ||
        child.name === 'weaponswap_left_arm' ||
        child.name === 'text_i_right' ||
        child.name === 'text_ii_right'
      ) {
        child.fields.rect.x = child.fields.rect.x + 227;
        child.fields.rect.y = child.fields.rect.y + 12;
      }
      return true;
    });
  D2RMM.writeJson(
    playerInventoryExpansionLayoutHDFilename,
    playerInventoryExpansionLayoutHD
  );

  const playerInventoryOriginalControllerLayoutHDFilename =
    'global\\ui\\layouts\\controller\\playerinventoryoriginallayouthd.json';
  const playerInventoryOriginalControllerLayoutHD = D2RMM.readJson(
    playerInventoryOriginalControllerLayoutHDFilename
  );
  playerInventoryOriginalControllerLayoutHD.children.forEach((child) => {
    if (child.name === 'background') {
      child.fields.filename =
        'Controller/Panel/InventoryPanel/V2/InventoryBG_Classic_Expanded';
      child.fields.rect.x = child.fields.rect.x - 166;
      child.fields.rect.y = child.fields.rect.y - 160;
    }
    if (child.name === 'grid') {
      child.fields.rect.x = child.fields.rect.x - 132;
      child.fields.rect.y = child.fields.rect.y - 344;
    }
    if (child.name === 'slot_right_arm') {
      child.fields.rect.x = child.fields.rect.x - 99;
      child.fields.rect.y = child.fields.rect.y - 60;
    }
    if (child.name === 'slot_left_arm') {
      child.fields.rect.x = child.fields.rect.x + 123;
      child.fields.rect.y = child.fields.rect.y - 62;
    }
    if (child.name === 'slot_torso') {
      child.fields.rect.x = child.fields.rect.x + 6;
      child.fields.rect.y = child.fields.rect.y - 199;
    }
    if (child.name === 'slot_head') {
      child.fields.rect.x = child.fields.rect.x - 239;
      child.fields.rect.y = child.fields.rect.y + 21;
    }
    if (child.name === 'slot_gloves') {
      child.fields.rect.x = child.fields.rect.x + 146;
      child.fields.rect.y = child.fields.rect.y - 282;
    }
    if (child.name === 'slot_feet') {
      child.fields.rect.x = child.fields.rect.x - 130;
      child.fields.rect.y = child.fields.rect.y - 281;
    }
    if (child.name === 'slot_belt') {
      child.fields.rect.x = child.fields.rect.x + 7;
      child.fields.rect.y = child.fields.rect.y - 185;
    }
    if (child.name === 'slot_neck') {
      child.fields.rect.x = child.fields.rect.x - 3;
      child.fields.rect.y = child.fields.rect.y - 167;
    }
    if (child.name === 'slot_right_hand') {
      child.fields.rect.x = child.fields.rect.x + 389;
      child.fields.rect.y = child.fields.rect.y - 417;
    }
    if (child.name === 'slot_left_hand') {
      child.fields.rect.x = child.fields.rect.x + 126;
      child.fields.rect.y = child.fields.rect.y - 417;
    }
    if (child.name === 'Belt') {
      child.fields.rect.x = child.fields.rect.x + 15;
      child.fields.rect.y = child.fields.rect.y + 595;
    }
    if (child.name === 'gold_amount' || child.name === 'gold_button') {
      child.fields.rect.x = child.fields.rect.x - 464;
      child.fields.rect.y = child.fields.rect.y + 20;
    }
  });
  D2RMM.writeJson(
    playerInventoryOriginalControllerLayoutHDFilename,
    playerInventoryOriginalControllerLayoutHD
  );

  const playerInventoryExpansionControllerLayoutHDFilename =
    'global\\ui\\layouts\\controller\\playerinventoryexpansionlayouthd.json';
  const playerInventoryExpansionControllerLayoutHD = D2RMM.readJson(
    playerInventoryExpansionControllerLayoutHDFilename
  );
  playerInventoryExpansionControllerLayoutHD.children.forEach((child) => {
    if (child.name === 'background') {
      child.fields.filename =
        'Controller/Panel/InventoryPanel/V2/InventoryBG_Expanded';
    }
    if (
      child.name === 'background_right_arm' ||
      child.name === 'background_right_arm_selected' ||
      child.name === 'WeaponSwapRightLegend' ||
      child.name === 'text_i_left' ||
      child.name === 'text_ii_left'
    ) {
      child.fields.rect.x = child.fields.rect.x - 99;
      child.fields.rect.y = child.fields.rect.y - 60;
    }
    if (
      child.name === 'background_left_arm' ||
      child.name === 'background_left_arm_selected' ||
      child.name === 'WeaponSwapLeftLegend' ||
      child.name === 'text_i_right' ||
      child.name === 'text_ii_right'
    ) {
      child.fields.rect.x = child.fields.rect.x + 123;
      child.fields.rect.y = child.fields.rect.y - 62;
    }
  });
  D2RMM.writeJson(
    playerInventoryExpansionControllerLayoutHDFilename,
    playerInventoryExpansionControllerLayoutHD
  );
}

D2RMM.copyFile(
  'hd', // <mod folder>\hd
  'hd', // <diablo 2 folder>\mods\<modname>\<modname>.mpq\data\hd
  true // overwrite any conflicts
);

// modify the stash save file to make sure it has 8 tab pages
if (config.isExtraTabsEnabled) {
  function getStashTabBinary(vers) {
    // prettier-ignore
    return [
      // extracted from first 68 bytes of an empty stash save file (SharedStashSoftCoreV2.d2i) of D2R v1.6.80273
      // the 9th byte seems to be related to the version of the game (e.g. 0x61, 0x62, 0x63) and a stash save
      // that uses a *newer* version than the currently running version of the game will fail to load
      0x55, 0xAA, 0x55, 0xAA, 0x01, 0x00, 0x00, 0x00, vers, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 16 bytes
      0x44, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 16 bytes
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 16 bytes
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // 16 bytes
      0x4A, 0x4D, 0x00, 0x00                                                                          //  4 bytes
    ];
  }

  function indexOf(haystack, needle, startIndex = 0) {
    let match = 0;
    const index = haystack.findIndex((value, index) => {
      if (index < startIndex) {
        return false;
      }
      // null matches any value
      if (needle[match] == null || value === needle[match]) {
        match++;
      } else {
        match = 0;
      }
      if (match === needle.length) {
        return true;
      }
      return false;
    });
    return index === -1 ? -1 : index - needle.length + 1;
  }

  function getStashTabStartIndices(stashData) {
    const stashTabPrefix = getStashTabBinary(null).slice(0, 10);
    const stashTabStartIndices = [];
    let index = -1;
    while (true) {
      index = indexOf(stashData, stashTabPrefix, index + 1);
      if (index !== -1) {
        stashTabStartIndices.push(index);
      } else {
        break;
      }
    }
    return stashTabStartIndices;
  }

  const results = {};
  function modSaveFile(filename) {
    const stashData = D2RMM.readSaveFile(filename);
    if (stashData == null) {
      console.debug(`Skipped ${filename} because the file was not found.`);
      results[filename] = false;
      return;
    }
    // backup existing stash tab data if it doesn't exist
    const stashDataBackup = D2RMM.readSaveFile(`${filename}.bak`);
    if (stashDataBackup == null) {
      D2RMM.writeSaveFile(`${filename}.bak`, stashData);
    }
    // find number of times prefix appears in stashData
    const stashTabIndices = getStashTabStartIndices(stashData);
    // find latest version code used by the save file
    const versionCode = Math.max(
      ...stashTabIndices.map((index) => stashData[index + 8])
    );
    // sanitize the data (each save files should have 3-7 shared tabs)
    const existingTabsCount = Math.max(3, Math.min(7, stashTabIndices.length));
    const tabsToAdd = 7 - existingTabsCount;
    // don't modify the save file if it doesn't need it
    if (tabsToAdd > 0) {
      D2RMM.writeSaveFile(
        filename,
        [].concat.apply(
          stashData,
          new Array(tabsToAdd).fill(getStashTabBinary(versionCode))
        )
      );
      console.debug(
        `Added ${tabsToAdd} additional shared stash tabs to ${filename} using version code 0x${versionCode.toString(
          16
        )}.`
      );
    } else {
      console.debug(
        `Skipped ${filename} because it already has 7 shared stash tabs.`
      );
    }
    results[filename] = true;
  }

  const SOFTCORE_SAVE_FILE = 'SharedStashSoftCoreV2.d2i';
  const HARDCORE_SAVE_FILE = 'SharedStashHardCoreV2.d2i';

  modSaveFile(SOFTCORE_SAVE_FILE);
  modSaveFile(HARDCORE_SAVE_FILE);

  if (!results[SOFTCORE_SAVE_FILE] && !results[HARDCORE_SAVE_FILE]) {
    console.warn(
      `Unable to enable additional shared stash tabs. Neither ${SOFTCORE_SAVE_FILE} nor ${HARDCORE_SAVE_FILE} were found.`
    );
  }
}
