if (D2RMM.getVersion == null || D2RMM.getVersion() < 1.6) {
  D2RMM.error('Requires D2RMM version 1.6 or higher.');
  return;
}

// ExpandedStash
{
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
}

// ExpandedInventory
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

// ExpandedCube
{
  const inventoryFilename = 'global\\excel\\inventory.txt';
  const inventory = D2RMM.readTsv(inventoryFilename);
  inventory.rows.forEach((row) => {
    if (
      row.class === 'Transmogrify Box Page 1' ||
      row.class === 'Transmogrify Box2'
    ) {
      row.gridX = 6;
    }
  });
  D2RMM.writeTsv(inventoryFilename, inventory);

  const profileHDFilename = 'global\\ui\\layouts\\_profilehd.json';
  const profileHD = D2RMM.readJson(profileHDFilename);
  D2RMM.writeJson(profileHDFilename, profileHD);

  const profileLVFilename = 'global\\ui\\layouts\\_profilelv.json';
  const profileLV = D2RMM.readJson(profileLVFilename);
  D2RMM.writeJson(profileLVFilename, profileLV);

  const horadricCubeLayoutFilename =
    'global\\ui\\layouts\\horadriccubelayout.json';
  const horadricCubeLayout = D2RMM.readJson(horadricCubeLayoutFilename);
  horadricCubeLayout.children.forEach((child) => {
    if (child.name === 'grid') {
      child.fields.cellCount.x = 6;
      child.fields.cellCount.y = 4;
      // TODO: shift left for new sprite
      // one cell = 29px, add 1 column to left and 2 columns to right of original space
      child.fields.rect.x = child.fields.rect.x - 29;
    }
    // TODO: new sprite
  });
  D2RMM.writeJson(horadricCubeLayoutFilename, horadricCubeLayout);

  const horadricCubeLayoutHDFilename =
    'global\\ui\\layouts\\horadriccubelayouthd.json';
  const horadricCubeHDLayout = D2RMM.readJson(horadricCubeLayoutHDFilename);
  horadricCubeHDLayout.children.forEach((child) => {
    if (child.name === 'grid') {
      child.fields.cellCount.x = 6;
      child.fields.cellCount.y = 4;
      child.fields.rect.x = child.fields.rect.x - 144;
    }
    if (child.name === 'background') {
      child.fields.filename = 'PANEL\\Horadric_Cube\\HoradricCube_BG_Expanded';
    }
  });
  D2RMM.writeJson(horadricCubeLayoutHDFilename, horadricCubeHDLayout);

  const controllerHoradricCubeHDLayoutFilename =
    'global\\ui\\layouts\\controller\\horadriccubelayouthd.json';
  const controllerHoradricCubeHDLayout = D2RMM.readJson(
    controllerHoradricCubeHDLayoutFilename
  );
  controllerHoradricCubeHDLayout.children.forEach((child) => {
    if (child.name === 'grid') {
      child.fields.rect.x = child.fields.rect.x - 142;
    }
    if (child.name === 'background') {
      child.fields.filename =
        'Controller/Panel/HoradricCube/V2/HoradricCubeBG_Expanded';
    }
  });
  D2RMM.writeJson(
    controllerHoradricCubeHDLayoutFilename,
    controllerHoradricCubeHDLayout
  );
}

// MercEquip
{
  const inventoryFilename = 'global\\excel\\inventory.txt';
  const inventory = D2RMM.readTsv(inventoryFilename);
  inventory.rows.forEach((row) => {
    if (row.class === 'Hireling') {
      row.rArmLeft = 340;
      row.rArmRight = 395;
      row.rArmTop = 27;
      row.rArmBottom = 139;
      row.torsoLeft = 453;
      row.torsoRight = 509;
      row.lArmLeft = 566;
      row.lArmRight = 621;
      row.lArmTop = 27;
      row.lArmBottom = 139;
      row.headLeft = 455;
      row.headRight = 509;
      row.neckLeft = 529;
      row.neckRight = 552;
      row.neckTop = 35;
      row.neckBottom = 59;
      row.neckWidth = 23;
      row.neckHeight = 24;
      row.rHandLeft = 415;
      row.rHandRight = 438;
      row.rHandTop = 180;
      row.rHandBottom = 204;
      row.rHandWidth = 23;
      row.rHandHeight = 24;
      row.lHandLeft = 529;
      row.lHandRight = 552;
      row.lHandTop = 180;
      row.lHandBottom = 204;
      row.lHandWidth = 23;
      row.lHandHeight = 24;
      row.beltLeft = 456;
      row.beltRight = 508;
      row.beltTop = 179;
      row.beltBottom = 204;
      row.beltWidth = 52;
      row.beltHeight = 25;
      row.feetLeft = 572;
      row.feetRight = 626;
      row.feetTop = 155;
      row.feetBottom = 207;
      row.feetWidth = 54;
      row.feetHeight = 52;
      row.glovesLeft = 341;
      row.glovesRight = 395;
      row.glovesTop = 153;
      row.glovesBottom = 205;
      row.glovesWidth = 54;
      row["glovesHeight\r"] = 53;
    } else if (row.class === 'Hireling2') {
      row.rArmLeft = 420;
      row.rArmRight = 475;
      row.rArmTop = 89;
      row.rArmBottom = 201;
      row.torsoLeft = 533;
      row.torsoRight = 589;
      row.lArmLeft = 646;
      row.lArmRight = 701;
      row.lArmTop = 89;
      row.lArmBottom = 201;
      row.headLeft = 535;
      row.headRight = 589;
      row.neckLeft = 609;
      row.neckRight = 632;
      row.neckTop = 95;
      row.neckBottom = 119;
      row.neckWidth = 23;
      row.neckHeight = 24;
      row.rHandLeft = 495;
      row.rHandRight = 518;
      row.rHandTop = 240;
      row.rHandBottom = 264;
      row.rHandWidth = 23;
      row.rHandHeight = 24;
      row.lHandLeft = 609;
      row.lHandRight = 632;
      row.lHandTop = 240;
      row.lHandBottom = 264;
      row.lHandWidth = 23;
      row.lHandHeight = 24;
      row.beltLeft = 536;
      row.beltRight = 588;
      row.beltTop = 239;
      row.beltBottom = 264;
      row.beltWidth = 52;
      row.beltHeight = 25;
      row.feetLeft = 652;
      row.feetRight = 706;
      row.feetTop = 217;
      row.feetBottom = 269;
      row.feetWidth = 54;
      row.feetHeight = 52;
      row.glovesLeft = 421;
      row.glovesRight = 475;
      row.glovesTop = 215;
      row.glovesBottom = 268;
      row.glovesWidth = 54;
      row["glovesHeight\r"] = 53;
    } else if (['Amazon', 'Sorceress', 'Necromancer', 'Paladin', 'Barbarian', 'Druid', 'Assassin']
      .indexOf(row.class) !== -1) {
      row.gridTop = 211;
      row.gridBottom = 441;
      row.rArmTop = 27;
      row.rArmBottom = 139;
      row.lArmLeft = 566;
      row.lArmRight = 621;
      row.lArmTop = 27;
      row.lArmBottom = 139;
      row.feetTop = 155;
      row.feetBottom = 207;
      row.glovesTop = 153;
      row.glovesBottom = 205;
    } else if (['Amazon2', 'Sorceress2', 'Necromancer2', 'Paladin2', 'Barbarian2', 'Druid2', 'Assassin2']
      .indexOf(row.class) !== -1) {
      row.gridTop = 271;
      row.gridBottom = 501;
      row.rArmTop = 89;
      row.rArmBottom = 201;
      row.lArmLeft = 646;
      row.lArmRight = 701;
      row.lArmTop = 89;
      row.lArmBottom = 201;
      row.feetTop = 217;
      row.feetBottom = 269;
      row.glovesTop = 215;
      row.glovesBottom = 268;
    } else if (row.class === 'Transmogrify Box Page 1') {
      row.gridLeft = 16;
      row.gridRight = 303;
      row.gridTop = 16;
      row.gridBottom = 254;
    } else if (row.class === 'Transmogrify Box2') {
      row.gridLeft = 97;
      row.gridRight = 382;
      row.gridTop = 75;
      row.gridBottom = 304;
    }
  });
  D2RMM.writeTsv(inventoryFilename, inventory);

  const itemTypesFilename = 'global\\excel\\itemtypes.txt';
  const itemtypes = D2RMM.readTsv(itemTypesFilename);
  itemtypes.rows.forEach((row) => {
    if (['Ring', 'Amulet', 'Boots', 'Gloves', 'Belt'].indexOf(row.ItemType) !== -1) {
      row.Equiv2 = 'merc';
    } else if (row.ItemType === 'Helm') {
      row.Code = 'merc';
      row.Equiv1 = '';
    } else if (row.ItemType === 'Expansion') {
      row["*eol\r"] = '0';
    }
  });
  const itemRecipe = {
    ItemType: 'Merc Equip',
    Code: 'helm',
    Equiv1: 'armo',
    Equiv2: 'merc',
    Repair: 1,
    Body: 1,
    BodyLoc1: 'head',
    BodyLoc2: 'head',
    Throwable: 0,
    Reload: 0,
    ReEquip: 0,
    AutoStack: 0,
    Rare: 1,
    Normal: 0,
    Beltable: 0,
    MaxSockets1: 2,
    MaxSocketsLevelThreshold1: 25,
    MaxSockets2: 2,
    MaxSocketsLevelThreshold2: 40,
    MaxSockets3: 3,
    TreasureClass: 0,
    Rarity: 3,
    VarInvGfx: 0,
    StorePage: 'armo',
    '*eol\r': 0,
  };
  itemtypes.rows.push({
    ...itemRecipe
  });
  D2RMM.writeTsv(itemTypesFilename, itemtypes);

  const hirelingHDFilename = 'global\\ui\\layouts\\hirelinginventorypanelhd.json';
  const hirelingHD = D2RMM.readJson(hirelingHDFilename);
  hirelingHD.children.forEach((child) => {
    if (child.type === 'ClickCatcherWidget') {
      child.fields.rect = {
        x: -180,
        y: -200,
        width: 1362,
        height: 1727
      };
    } else if (child.type === 'TextBoxWidget') {
      if (child.name === 'Title') {
        child.fields.rect = {
          x: 481,
          y: -69,
          width: 196,
          height: 196
        };
      } else if (child.name === 'CharacterName') {
        child.fields.rect.x = 121;
        child.fields.rect.y = 849;
        child.fields.style.alignment.v = 'bottom';
      } else if (child.name === 'HPTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1020;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'HPStat') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1020;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'HireTypeText') {
        child.fields.rect.x = 126;
        child.fields.rect.y = 937;
        child.fields.rect.width = 100;
        child.fields.rect.height = 30;
        child.fields.style.pointSize = '$XMediumFontSize';
        delete child.fields.style.alignment.v;
      } else if (child.name === 'StrengthTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1102;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'StrengthStat') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1102;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DexTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1186;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DexStat') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1186;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DamageTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DamageStat') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ArmorClassTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1353;
        child.fields.rect.width = 251;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ArmorClass') {
        child.fields.rect.x = 361;
        child.fields.rect.y = 1353;
        child.fields.rect.width = 187;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'FireResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1103;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'FireText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1103;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ColdResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1184;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ColdText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1184;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'LightningResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'LightningText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'PoisonResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1351;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'PoisonText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1351;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      }
    } else if (child.type === 'InventorySlotWidget') {
      if (child.name === 'slot_head') {
        child.fields.rect.y = 113;
        delete child.fields['gemSocketFilename'];
      } else if (child.name === 'slot_torso') {
        child.fields.rect.x = 483;
        child.fields.rect.y = 348;
        delete child.fields['gemSocketFilename'];
      } else if (child.name === 'slot_right_arm') {
        child.fields.rect.x = 110;
        child.fields.rect.y = 156;
        child.fields.location = 'right_arm';
      } else if (child.name === 'slot_left_arm') {
        child.fields.rect.x = 863;
        child.fields.rect.y = 156;
        child.fields.location = 'left_arm';
      }
    } else if (child.type === 'ExpBarWidget') {
      child.fields.rect.y = 913;
    } else if (child.type === 'ButtonWidget') {
      if (child.name === 'AdvancedStats') {
        child.fields.rect.x = 1104;
        child.fields.rect.y = 713;
        delete child.fields['hoveredFrame'];
        delete child.fields['pressLabelOffset'];
      } else if (child.name === 'CloseButton') {
        delete child.fields['sound'];
      }
    } else if (child.type === 'HirelingSkillIconWidget') {
      if (child.name === 'Skill0') {
        child.fields.rect.x = 673;
      } else if (child.name === 'Skill1') {
        child.fields.rect.x = 780;
      } else if (child.name === 'Skill2') {
        child.fields.rect.x = 887;
      }
      child.fields.rect.y = 1007;
      child.fields.rect.scale = 0.60;
    } else if (child.type === 'Widget') {
      if (child.name === 'Damage') {
        child.fields.rect.y = 1269;
      }
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_belt",
    fields: {
      rect: {
        x: 483,
        y: 688,
        width: 196,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "belt",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Belt",
      isHireable: true
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_right_hand",
    fields: {
      rect: {
        x: 720,
        y: 691,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "right_hand",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Ring",
      isHireable: false
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_left_hand",
    fields: {
      rect: {
        x: 344,
        y: 691,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "left_hand",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Ring",
      isHireable: false
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_gloves",
    fields: {
      rect: {
        x: 108,
        y: 591,
        width: 196,
        height: 196
      },
      cellSize: "$ItemCellSize",
      location: "gloves",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Glove",
      isHireable: false
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_feet",
    fields: {
      rect: {
        x: 861,
        y: 591,
        width: 196,
        height: 196
      },
      cellSize: "$ItemCellSize",
      location: "feet",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Boots",
      isHireable: false
    }
  });
  hirelingHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_neck",
    fields: {
      rect: {
        x: 720,
        y: 271,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "neck",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Amulet",
      isHireable: false
    }
  });
  D2RMM.writeJson(hirelingHDFilename, hirelingHD);

  const profileCHDFilename = 'global\\ui\\layouts\\controller\\_profilehd.json';
  const profileCHD = D2RMM.readJson(profileCHDFilename);
  profileCHD.ConsoleLeftPanelAnchor = { "x": 0.521, "y": 0.387 };
  D2RMM.writeJson(profileCHDFilename, profileCHD);

  const hirelingCHDFilename = 'global\\ui\\layouts\\controller\\hirelinginventorypanelhd.json';
  const hirelingCHD = D2RMM.readJson(hirelingCHDFilename);
  delete hirelingCHD['basedOn'];
  hirelingCHD.fields = {
    rect: "$ConsoleLeftPanelRect",
    anchor: "$ConsoleLeftPanelAnchor",
    bowBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_Bow",
    spearBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_Spear",
    longswordBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_LongSword",
    shieldBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_Shield",
    twoHandSwordBackgroundFilename: "PANEL\\Hireling\\HireablePanel\\Hireables_Paperdoll_2HSword"
  };
  hirelingCHD.children.forEach((child) => {
    if (child.type === 'ImageWidget') {
      child.fields = {
        rect: {
          x: 0,
          y: 0
        },
        filename: "\\PANEL\\Hireling\\HirelingPanel"
      }
    } else if (child.type === 'ClickCatcherWidget') {
      child.fields = {
        rect: {
          x: -180,
          y: -200,
          width: 1362,
          height: 1727
        }
      };
    } else if (child.type === 'InventorySlotWidget') {
      if (child.name === 'slot_head') {
        child.fields.rect.x = 481;
        child.fields.rect.y = 113;
        child.fields.location = 'head';
        child.fields.backgroundFilename = 'PANEL\\Inventory\\Inventory_Paperdoll_HeadArmor';
      } else if (child.name === 'slot_torso') {
        child.fields.rect.x = 483;
        child.fields.rect.y = 348;
        child.fields.location = 'torso';
        child.fields.backgroundFilename = 'PANEL\\Inventory\\Inventory_Paperdoll_ChestArmor';
      } else if (child.name === 'slot_right_arm') {
        child.fields.rect.x = 110;
        child.fields.rect.y = 156;
        child.fields.location = 'right_arm';
        child.fields.gemSocketFilename = 'PANEL\\gemsocket';
        child.fields.backgroundFilename = 'PANEL\\Inventory\\Inventory_Paperdoll_Weapon';
      } else if (child.name === 'slot_left_arm') {
        child.fields.rect.x = 863;
        child.fields.rect.y = 156;
        child.fields.location = 'left_arm';
        child.fields.gemSocketFilename = 'PANEL\\gemsocket';
        child.fields.backgroundFilename = 'PANEL\\Inventory\\Inventory_Paperdoll_Weapon';
      }
      child.fields.cellSize = '$ItemCellSize';
      child.fields.isHireable = true;
      delete child.fields['navigation'];
    } else if (child.type === 'TextBoxWidget') {
      if (child.name === 'CharacterName') {
        child.fields.rect.x = 121;
        child.fields.rect.y = 849;
        child.fields.rect.width = 500;
        child.fields.rect.height = 50;
        child.fields.style.fontColor = '$FontColorWhite';
        child.fields.style.pointSize = '$LargeFontSize';
        child.fields.style.dropShadow = '$DefaultDropShadow';
      } else if (child.name === 'HPTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1020;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrlif';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'HPStat') {
        child.fields.rect.y = 1020;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'HireTypeText') {
        child.fields.rect.x = 126;
        child.fields.rect.y = 937;
        child.fields.rect.width = 100;
        child.fields.rect.height = 30;
        child.fields.text = '@strchrlvl';
        child.fields.style.fontColor = '$FontColorWhite';
        child.fields.style.pointSize = '$XMediumFontSize';
        child.fields.style.dropShadow = '$DefaultDropShadow';
        delete child.fields.style.alignment.v;
      } else if (child.name === 'StrengthTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1102;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrstr';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'StrengthStat') {
        child.fields.rect.y = 1102;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DexTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1186;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrdex';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DexStat') {
        child.fields.rect.y = 1186;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DamageTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrskm';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'DamageStat') {
        child.fields.rect.y = 1269;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ArmorClassTitle') {
        child.fields.rect.x = 104;
        child.fields.rect.y = 1353;
        child.fields.rect.width = 251;
        child.fields.rect.height = 59;
        child.fields.text = '@strchrdef';
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ArmorClass') {
        child.fields.rect.y = 1353;
        child.fields.rect.width = 187;
        child.fields.rect.height = 59;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'FireResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1103;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
        delete child.fields['useAltStyleIfDoesntFit'];
        delete child.fields['altStyle'];
      } else if (child.name === 'FireText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1103;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'ColdResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1184;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
        delete child.fields['useAltStyleIfDoesntFit'];
        delete child.fields['altStyle'];
      } else if (child.name === 'ColdText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1184;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'LightningResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
        delete child.fields['useAltStyleIfDoesntFit'];
        delete child.fields['altStyle'];
      } else if (child.name === 'LightningText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1269;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      } else if (child.name === 'PoisonResistanceTitle') {
        child.fields.rect.x = 608;
        child.fields.rect.y = 1351;
        child.fields.rect.width = 325;
        child.fields.style.pointSize = '$SmallFontSize';
        delete child.fields['useAltStyleIfDoesntFit'];
        delete child.fields['altStyle'];
      } else if (child.name === 'PoisonText') {
        child.fields.rect.x = 938;
        child.fields.rect.y = 1351;
        child.fields.rect.width = 113;
        child.fields.style.pointSize = '$SmallFontSize';
      }
    } else if (child.type === 'ExpBarWidget') {
      child.fields = {
        rect: {
          x: 139,
          y: 913,
          width: 888,
          height: 10
        },
        filename: "PANEL\\Hireling\\Hireling_ExpBar",
        frame: 0,
        isHireling: true,
        hitMargin: {
          top: 15,
          bottom: 15
        }
      }
    } else if (child.type === 'HirelingSkillIconWidget') {
      if (child.name === 'Skill0') {
        child.fields.rect.x = 673;
      } else if (child.name === 'Skill1') {
        child.fields.rect.x = 780;
      } else if (child.name === 'Skill2') {
        child.fields.rect.x = 887;
      }

      child.fields.rect.y = 1007;
      child.fields.rect.scale = 0.60;
    } else if (child.type === 'Widget') {
      if (child.name === 'Damage') {
        child.fields = {
          rect: {
            x: 328,
            y: 1269,
            width: 237,
            height: 59
          }
        };
        child.children.forEach((wdChild) => {
          wdChild.fields.rect.width = 237;
          wdChild.fields.rect.height = 59;
          wdChild.fields.style.pointSize = '$SmallPanelFontSize';
          if (wdChild.name === 'DamageStatTop') {
            wdChild.fields.rect.y = 0;
          } else if (wdChild.name === 'DamageStatBottom') {
            wdChild.fields.rect.y = -1;
          }
        });
      }
    }
  });
  hirelingCHD.children.push({
    type: "ImageWidget",
    name: "LeftHinge",
    fields: {
      rect: "$LeftHingeRect",
      filename: "$LeftHingeSprite"
    }
  });
  hirelingCHD.children.push({
    type: "TextBoxWidget",
    name: "Title",
    fields: {
      rect: {
        x: 481,
        y: -69,
        width: 196,
        height: 196
      },
      style: "$StyleTitleBlock",
      text: "@MiniPanelHireinv"
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_belt",
    fields: {
      rect: {
        x: 483,
        y: 688,
        width: 196,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "belt",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Belt",
      isHireable: true
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_right_hand",
    fields: {
      rect: {
        x: 720,
        y: 691,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "right_hand",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Ring",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_left_hand",
    fields: {
      rect: {
        x: 344,
        y: 691,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "left_hand",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Ring",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_gloves",
    fields: {
      rect: {
        x: 108,
        y: 591,
        width: 196,
        height: 196
      },
      cellSize: "$ItemCellSize",
      location: "gloves",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Glove",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_feet",
    fields: {
      rect: {
        x: 861,
        y: 591,
        width: 196,
        height: 196
      },
      cellSize: "$ItemCellSize",
      location: "feet",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Boots",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "InventorySlotWidget",
    name: "slot_neck",
    fields: {
      rect: {
        x: 720,
        y: 271,
        width: 98,
        height: 98
      },
      cellSize: "$ItemCellSize",
      location: "neck",
      gemSocketFilename: "PANEL\\gemsocket",
      backgroundFilename: "PANEL\\Inventory\\Inventory_Paperdoll_Amulet",
      isHireable: false
    }
  });
  hirelingCHD.children.push({
    type: "ButtonWidget",
    name: "AdvancedStats",
    fields: {
      rect: {
        x: 1104,
        y: 713
      },
      filename: "PANEL\\Character_Sheet\\AdvancedStatsButton",
      onClickMessage: "HirelingInventoryPanelMessage:ToggleAdvancedStats",
      tooltipString: "@AdvancedStats"
    }
  });
  hirelingCHD.children.push({
    type: "ButtonWidget",
    name: "CloseButton",
    fields: {
      rect: {
        x: 1075,
        y: 9
      },
      filename: "PANEL\\closebtn_4x",
      hoveredFrame: 3,
      onClickMessage: "HirelingInventoryPanelMessage:Close",
      tooltipString: "@strClose"
    }
  });
  D2RMM.writeJson(hirelingCHDFilename, hirelingCHD);
}

// StackableGems
{
  const SINGLE_ITEM_CODE = 'gem';
  const STACK_ITEM_CODE = 'sgem';

  const ITEM_TYPES = [];
  function converItemTypeToStackItemType(itemtype) {
    if (itemtype != null && ITEM_TYPES.indexOf(itemtype) !== -1) {
      // original idea to use "z" as prefix ran into issues due to "zlb" already being taken
      // in favor of backwards compatibility, only changing this one conflict
      const prefix = itemtype === 'glb' ? 'q' : 'z';
      return `${prefix}${itemtype.slice(1)}`;
    }
    return itemtype;
  }

  const miscFilenames = [];

  const itemsFilename = 'hd\\items\\items.json';
  const items = D2RMM.readJson(itemsFilename);
  const newItems = [...items];
  for (const index in items) {
    const item = items[index];
    for (const itemtype in item) {
      const asset = item[itemtype].asset;
      if (
        asset.startsWith(`${SINGLE_ITEM_CODE}/`) &&
        !asset.endsWith('_stack') &&
        itemtype !== 'jew' // exclude jewels
      ) {
        ITEM_TYPES.push(itemtype);
        const itemtypeStack = converItemTypeToStackItemType(itemtype);
        newItems.push({ [itemtypeStack]: { asset: `${asset}_stack` } });
        miscFilenames.push(asset.replace(`${SINGLE_ITEM_CODE}/`, ''));
      }
    }
  }
  D2RMM.writeJson(itemsFilename, newItems);

  const miscDirFilename = `hd\\items\\misc\\${SINGLE_ITEM_CODE}\\`;
  for (const index in miscFilenames) {
    const miscFilename = `${miscDirFilename + miscFilenames[index]}.json`;
    const miscStackFilename = `${miscDirFilename + miscFilenames[index]
      }_stack.json`;
    const miscStack = D2RMM.readJson(miscFilename);
    D2RMM.writeJson(miscStackFilename, miscStack);
  }

  const itemtypesFilename = 'global\\excel\\itemtypes.txt';
  const itemtypes = D2RMM.readTsv(itemtypesFilename);
  itemtypes.rows.forEach((itemtype) => {
    if (itemtype.Code === SINGLE_ITEM_CODE) {
      itemtypes.rows.push({
        ...itemtype,
        ItemType: `${itemtype.ItemType} Stack`,
        Code: STACK_ITEM_CODE,
        Equiv1: 'misc',
        AutoStack: 1,
      });
    }
  });
  D2RMM.writeTsv(itemtypesFilename, itemtypes);

  if (config.default) {
    const treasureclassexFilename = 'global\\excel\\treasureclassex.txt';
    const treasureclassex = D2RMM.readTsv(treasureclassexFilename);
    treasureclassex.rows.forEach((treasureclass) => {
      treasureclass.Item1 = converItemTypeToStackItemType(treasureclass.Item1);
      treasureclass.Item2 = converItemTypeToStackItemType(treasureclass.Item2);
      treasureclass.Item3 = converItemTypeToStackItemType(treasureclass.Item3);
      treasureclass.Item4 = converItemTypeToStackItemType(treasureclass.Item4);
      treasureclass.Item5 = converItemTypeToStackItemType(treasureclass.Item5);
      treasureclass.Item6 = converItemTypeToStackItemType(treasureclass.Item6);
      treasureclass.Item7 = converItemTypeToStackItemType(treasureclass.Item7);
    });
    D2RMM.writeTsv(treasureclassexFilename, treasureclassex);
  }

  const miscFilename = 'global\\excel\\misc.txt';
  const misc = D2RMM.readTsv(miscFilename);
  misc.rows.forEach((item) => {
    if (ITEM_TYPES.indexOf(item.code) !== -1) {
      const itemStack = {
        ...item,
        name: `${item.name} Stack`,
        compactsave: 0,
        type: STACK_ITEM_CODE,
        code: converItemTypeToStackItemType(item.code),
        stackable: 1,
        minstack: 1,
        maxstack: config.maxStack,
        spawnstack: 1,
        spelldesc: 2,
        spelldescstr: 'StackableGem',
        spelldesccolor: 0,
      };
      delete itemStack.type2;
      misc.rows.push(itemStack);
      item.spawnable = 0;
    }
  });
  D2RMM.writeTsv(miscFilename, misc);

  const itemNamesFilename = 'local\\lng\\strings\\item-names.json';
  const itemNames = D2RMM.readJson(itemNamesFilename);
  ITEM_TYPES.forEach((itemtype) => {
    const itemName = itemNames.find(({ Key }) => Key === itemtype);
    if (itemName != null) {
      const stacktype = converItemTypeToStackItemType(itemtype);
      itemNames.push({
        ...itemName,
        id: D2RMM.getNextStringID(),
        Key: stacktype,
      });
    }
  });
  D2RMM.writeJson(itemNamesFilename, itemNames);

  const itemModifiersFilename = 'local\\lng\\strings\\item-modifiers.json';
  const itemModifiers = D2RMM.readJson(itemModifiersFilename);
  itemModifiers.push({
    id: D2RMM.getNextStringID(),
    Key: 'StackableGem',
    enUS: 'Can be transmuted into a usable gem',
    zhTW: '可以转化为可用的宝石',
    deDE: 'Kann in einen nutzbaren Edelstein umgewandelt werden',
    esES: 'Se puede transmutar en una gema utilizable',
    frFR: 'Peut être transmuté en gemme utilisable',
    itIT: 'Può essere trasmutato in una gemma utilizzabile',
    koKR: '사용 가능한 보석으로 변환할 수 있습니다',
    plPL: 'Może zostać przekształcony w użyteczny klejnot',
    esMX: 'Se puede transmutar en una gema utilizable',
    jaJP: '使用可能な宝石に変換することができます',
    ptBR: 'Pode ser transmutado em uma gema utilizável',
    ruRU: 'Может быть превращен в полезный драгоценный камень',
    zhCN: '可以转化为可用的宝石',
  });
  D2RMM.writeJson(itemModifiersFilename, itemModifiers);

  const cubemainFilename = 'global\\excel\\cubemain.txt';
  const cubemain = D2RMM.readTsv(cubemainFilename);
  for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
    const itemtype = ITEM_TYPES[i];
    const stacktype = converItemTypeToStackItemType(itemtype);
    // convert from single to stack
    cubemain.rows.push({
      description: `${itemtype} -> ${stacktype}`,
      enabled: 1,
      version: 100,
      numinputs: 1,
      'input 1': itemtype,
      output: stacktype,
      '*eol': 0,
    });
    // convert from stack to single
    cubemain.rows.push({
      description: `${stacktype} -> ${itemtype}`,
      enabled: 1,
      version: 100,
      op: 18, // skip recipe if item's Stat.Accr(param) != value
      param: 70, // quantity (itemstatcost.txt)
      value: 1, // only execute rule if quantity == 1
      numinputs: 1,
      'input 1': stacktype,
      output: itemtype,
      '*eol': 0,
    });
  }

  // this is behind a config option because it's *a lot* of recipes...
  if (config.convertWhenDestacking) {
    for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
      const itemtype = ITEM_TYPES[i];
      const stacktype = converItemTypeToStackItemType(itemtype);
      for (let j = 2; j <= config.maxStack; j = j + 1) {
        cubemain.rows.push({
          description: `Stack of ${j} ${itemtype} -> Stack of ${j - 1
            } and Stack of 1`,
          enabled: 1,
          version: 0,
          op: 18, // skip recipe if item's Stat.Accr(param) != value
          param: 70, // quantity (itemstatcost.txt)
          value: j, // only execute rule if quantity == j
          numinputs: 1,
          'input 1': stacktype,
          output: `"${stacktype},qty=${j - 1}"`,
          'output b': `"${itemtype},qty=1"`,
          '*eol': 0,
        });
      }
    }
  }
  // if another mod already added destacking, don't add it again
  else if (
    cubemain.rows.find(
      (row) => row.description === 'Stack of 2 -> Stack of 1 and Stack of 1'
    ) == null
  ) {
    for (let i = 2; i <= config.maxStack; i = i + 1) {
      cubemain.rows.push({
        description: `Stack of ${i} -> Stack of ${i - 1} and Stack of 1`,
        enabled: 1,
        version: 0,
        op: 18, // skip recipe if item's Stat.Accr(param) != value
        param: 70, // quantity (itemstatcost.txt)
        value: i, // only execute rule if quantity == i
        numinputs: 1,
        'input 1': 'misc',
        output: `"usetype,qty=${i - 1}"`,
        'output b': `"usetype,qty=1"`,
        '*eol': 0,
      });
    }
  }

  if (config.bulkUpgrade) {
    for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
      // no upgrade for perfect gems
      if ((i + 1) % 5 == 0) {
        continue;
      }
      const itemtype = ITEM_TYPES[i];
      const stacktype = converItemTypeToStackItemType(itemtype);
      const upgradedItemtype = ITEM_TYPES[i + 1];
      const upgradedStacktype = converItemTypeToStackItemType(upgradedItemtype);
      for (let j = 30; j < config.maxStack; j = j + 1) {
        cubemain.rows.push({
          description:
            `Stack of ${j} ${itemtype} + 1 id scroll -> Stack` +
            ` of 10 ${upgradedItemtype} + Stack of ${j - 30} ${itemtype} + 1` +
            ` id scroll`,
          enabled: 1,
          version: 0,
          op: 18, // skip recipe if item's Stat.Accr(param) != value
          param: 70, // quantity (itemstatcost.txt)
          value: j, // only execute rule if quantity == j
          numinputs: 2,
          'input 1': stacktype,
          'input 2': 'isc',
          output: `"${upgradedStacktype},qty=10"`,
          'output b': j == 30 ? null : `"${stacktype},qty=${j - 30}"`,
          'output c': 'isc',
          '*eol': 0,
        });
      }
    }
  }
  D2RMM.writeTsv(cubemainFilename, cubemain);
}

// StackableRunes
{
  const SINGLE_ITEM_CODE = 'rune';
  const STACK_ITEM_CODE = 'runs';

  const ITEM_TYPES = [];
  function converItemTypeToStackItemType(itemtype) {
    if (itemtype != null && ITEM_TYPES.indexOf(itemtype) !== -1) {
      return itemtype.replace(/^r/, 's');
    }
    return itemtype;
  }

  const miscFilenames = [];

  const itemsFilename = 'hd\\items\\items.json';
  const items = D2RMM.readJson(itemsFilename);
  const newItems = [...items];
  for (const index in items) {
    const item = items[index];
    for (const itemtype in item) {
      const asset = item[itemtype].asset;
      if (asset.startsWith(`${SINGLE_ITEM_CODE}/`) && !asset.endsWith('_stack')) {
        ITEM_TYPES.push(itemtype);
        const itemtypeStack = converItemTypeToStackItemType(itemtype);
        newItems.push({ [itemtypeStack]: { asset: `${asset}_stack` } });
        miscFilenames.push(asset.replace(`${SINGLE_ITEM_CODE}/`, ''));
      }
    }
  }
  D2RMM.writeJson(itemsFilename, newItems);

  const miscDirFilename = `hd\\items\\misc\\${SINGLE_ITEM_CODE}\\`;
  for (const index in miscFilenames) {
    const miscFilename = `${miscDirFilename + miscFilenames[index]}.json`;
    const miscStackFilename = `${miscDirFilename + miscFilenames[index]
      }_stack.json`;
    const miscStack = D2RMM.readJson(miscFilename);
    D2RMM.writeJson(miscStackFilename, miscStack);
  }

  const itemtypesFilename = 'global\\excel\\itemtypes.txt';
  const itemtypes = D2RMM.readTsv(itemtypesFilename);
  itemtypes.rows.forEach((itemtype) => {
    if (itemtype.Code === SINGLE_ITEM_CODE) {
      itemtypes.rows.push({
        ...itemtype,
        ItemType: `${itemtype.ItemType} Stack`,
        Code: STACK_ITEM_CODE,
        Equiv1: 'misc',
        AutoStack: 1,
      });
    }
  });
  D2RMM.writeTsv(itemtypesFilename, itemtypes);

  if (config.default) {
    const treasureclassexFilename = 'global\\excel\\treasureclassex.txt';
    const treasureclassex = D2RMM.readTsv(treasureclassexFilename);
    treasureclassex.rows.forEach((treasureclass) => {
      treasureclass.Item1 = converItemTypeToStackItemType(treasureclass.Item1);
      treasureclass.Item2 = converItemTypeToStackItemType(treasureclass.Item2);
    });
    D2RMM.writeTsv(treasureclassexFilename, treasureclassex);
  }

  const miscFilename = 'global\\excel\\misc.txt';
  const misc = D2RMM.readTsv(miscFilename);
  misc.rows.forEach((item) => {
    if (ITEM_TYPES.indexOf(item.code) !== -1) {
      misc.rows.push({
        ...item,
        name: `${item.name} Stack`,
        compactsave: 0,
        type: STACK_ITEM_CODE,
        code: converItemTypeToStackItemType(item.code),
        stackable: 1,
        minstack: 1,
        maxstack: config.maxStack,
        spawnstack: 1,
        spelldesc: 2,
        spelldescstr: 'StackableRune',
        spelldesccolor: 0,
      });
      item.spawnable = 0;
    }
  });
  D2RMM.writeTsv(miscFilename, misc);

  const itemNamesFilename = 'local\\lng\\strings\\item-names.json';
  const itemNames = D2RMM.readJson(itemNamesFilename);
  ITEM_TYPES.forEach((itemtype) => {
    const itemName = itemNames.find(({ Key }) => Key === itemtype);
    if (itemName != null) {
      const stacktype = converItemTypeToStackItemType(itemtype);
      itemNames.push({
        ...itemName,
        id: D2RMM.getNextStringID(),
        Key: stacktype,
      });
    }
  });
  D2RMM.writeJson(itemNamesFilename, itemNames);

  const itemModifiersFilename = 'local\\lng\\strings\\item-modifiers.json';
  const itemModifiers = D2RMM.readJson(itemModifiersFilename);
  itemModifiers.push({
    id: D2RMM.getNextStringID(),
    Key: 'StackableRune',
    enUS: 'Can be transmuted into a usable rune',
    zhTW: '可以转化为可用的符文',
    deDE: 'Kann in eine nutzbare Rune umgewandelt werden',
    esES: 'Se puede transmutar en una runa utilizable',
    frFR: 'Peut être transmuté en une rune utilisable',
    itIT: 'Può essere trasmutato in una runa utilizzabile',
    koKR: '사용 가능한 룬으로 변환할 수 있습니다',
    plPL: 'Może zostać przemieniony w użyteczną runę',
    esMX: 'Se puede transmutar en una runa utilizable',
    jaJP: '使用可能なルーンに変換できます',
    ptBR: 'Pode ser transmutado em uma runa utilizável',
    ruRU: 'Может быть преобразован в полезную руну',
    zhCN: '可以转化为可用的符文',
  });
  D2RMM.writeJson(itemModifiersFilename, itemModifiers);

  const cubemainFilename = 'global\\excel\\cubemain.txt';
  const cubemain = D2RMM.readTsv(cubemainFilename);
  for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
    const itemtype = ITEM_TYPES[i];
    const stacktype = converItemTypeToStackItemType(itemtype);
    // convert from single to stack
    cubemain.rows.push({
      description: `${itemtype} -> ${stacktype}`,
      enabled: 1,
      version: 100,
      numinputs: 1,
      'input 1': itemtype,
      output: stacktype,
      '*eol': 0,
    });
    // convert from stack to single
    cubemain.rows.push({
      description: `${stacktype} -> ${itemtype}`,
      enabled: 1,
      version: 100,
      op: 18, // skip recipe if item's Stat.Accr(param) != value
      param: 70, // quantity (itemstatcost.txt)
      value: 1, // only execute rule if quantity == 1
      numinputs: 1,
      'input 1': stacktype,
      output: itemtype,
      '*eol': 0,
    });
  }

  // this is behind a config option because it's *a lot* of recipes...
  if (config.convertWhenDestacking) {
    for (let i = 0; i < ITEM_TYPES.length; i = i + 1) {
      const itemtype = ITEM_TYPES[i];
      const stacktype = converItemTypeToStackItemType(itemtype);
      for (let j = 2; j <= config.maxStack; j = j + 1) {
        cubemain.rows.push({
          description: `Stack of ${j} ${itemtype} -> Stack of ${j - 1
            } and Stack of 1`,
          enabled: 1,
          version: 0,
          op: 18, // skip recipe if item's Stat.Accr(param) != value
          param: 70, // quantity (itemstatcost.txt)
          value: j, // only execute rule if quantity == j
          numinputs: 1,
          'input 1': stacktype,
          output: `"${stacktype},qty=${j - 1}"`,
          'output b': `"${itemtype},qty=1"`,
          '*eol': 0,
        });
      }
    }
  }
  // if another mod already added destacking, don't add it again
  else if (
    cubemain.rows.find(
      (row) => row.description === 'Stack of 2 -> Stack of 1 and Stack of 1'
    ) == null
  ) {
    for (let i = 2; i <= config.maxStack; i = i + 1) {
      cubemain.rows.push({
        description: `Stack of ${i} -> Stack of ${i - 1} and Stack of 1`,
        enabled: 1,
        version: 0,
        op: 18, // skip recipe if item's Stat.Accr(param) != value
        param: 70, // quantity (itemstatcost.txt)
        value: i, // only execute rule if quantity == i
        numinputs: 1,
        'input 1': 'misc',
        output: `"usetype,qty=${i - 1}"`,
        'output b': `"usetype,qty=1"`,
        '*eol': 0,
      });
    }
  }
  D2RMM.writeTsv(cubemainFilename, cubemain);

  // D2R colors runes as orange by default, but it seems to be based on item type
  // rather than localization strings so it does not apply to the stacked versions
  // we update the localization file to manually color the names of runes here
  // so that it will also apply to the stacked versions of the runes
  const itemRunesFilename = 'local\\lng\\strings\\item-runes.json';
  const itemRunes = D2RMM.readJson(itemRunesFilename);
  itemRunes.forEach((item) => {
    const itemtype = item.Key;
    if (itemtype.match(/^r[0-9]{2}$/) != null) {
      // update all localizations
      for (const key in item) {
        if (key !== 'id' && key !== 'Key') {
          if (item[key].indexOf('ÿc') !== -1) {
            // if the rune is already colored by something else (e.g. another mod)
            // then respect that change and don't override it
            continue;
          }
          // no idea what this is, but color codes before [fs] don't work
          const [, prefix = '', value] = item[key].match(/^(\[fs\])?(.*)$/) ?? [
            '',
            '',
            item[key],
          ];
          item[key] = `${prefix}ÿc8${value}`;
        }
      }
    }
  });
  D2RMM.writeJson(itemRunesFilename, itemRunes);
}

// Improved rune drop and general drop rate
{
  function SetDefaultQuestProp(row, setNoDrop) {
    row.Unique = 1024;
    row.Set = 800;
    row.Rare = 800;
    row.Magic = 800;
    if (setNoDrop) {
      row.NoDrop = 0;
    }
  }
  function SetQuestProp(row, pOne) {
    SetDefaultQuestProp(row, true);
    row.Item1 = 'gld,mul=2000';
    row.Prob1 = pOne;
  }
  const treasureclassexFilename = 'global\\excel\\treasureclassex.txt';
  const treasureclassex = D2RMM.readTsv(treasureclassexFilename);
  treasureclassex.rows.forEach((row) => {
    const treasureClass = row['Treasure Class'];
    // Improved rune drop
    {
      if (treasureClass === 'Runes 17') {
        row.Prob1 = 3;
        row.Prob2 = 16;
      }
      if (treasureClass === 'Runes 16') {
        row.Prob3 = 15;
        row.Prob4 = 7;
      }
      if (treasureClass === 'Runes 15') {
        row.Prob3 = 14;
        row.Prob4 = 7;
      }
      if (treasureClass === 'Runes 14') {
        row.Prob3 = 13;
        row.Prob4 = 8;
      }
      if (treasureClass === 'Runes 13') {
        row.Prob3 = 12;
        row.Prob4 = 8;
      }
      if (treasureClass === 'Runes 12') {
        row.Prob3 = 11;
        row.Prob4 = 4;
      }
      if (treasureClass === 'Runes 11') {
        row.Prob3 = 10;
      }
      if (treasureClass === 'Runes 10') {
        row.Prob3 = 9;
      }
      if (treasureClass === 'Runes 9') {
        row.Prob3 = 8;
      }
      if (treasureClass === 'Runes 8') {
        row.Prob3 = 7;
      }
      if (treasureClass === 'Runes 7') {
        row.Prob3 = 6;
      }
      if (treasureClass === 'Runes 6') {
        row.Prob3 = 5;
      }
      if (treasureClass === 'Runes 5') {
        row.Prob3 = 4;
      }
      if (treasureClass === 'Runes 4') {
        row.Prob3 = 3;
      }
      if (treasureClass === 'Runes 3') {
        row.Prob3 = 3;
      }
      if (treasureClass === 'Runes 2') {
        row.Prob3 = 3;
      }
    }

    // Increased general drop rate from Bosses
    {
      if (treasureClass === 'Andariel') {
        SetQuestProp(row, 5);
        row.Prob2 = 15;
        row.Item3 = 'Act 2 Equip C';
        row.Prob3 = 5;
        row.Item4 = 'rin';
        row.Prob4 = 2;
      }
      if (treasureClass === 'Andariel (N)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=3536';
        row.Prob1 = 5;
        row.Prob2 = 19;
        row.Item3 = 'Act 2 (N) Equip C';
        row.Prob3 = 6;
        row.Item4 = 'Act 2 (N) Good';
        row.Prob4 = 3;
      }
      if (treasureClass === 'Andariel (H)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=4048';
        row.Prob1 = 5;
        row.Item3 = 'Act 2 (H) Equip C';
        row.Prob3 = 7;
        row.Prob4 = 5;
        row.Prob5 = 0;
      }
      if (treasureClass === 'Andarielq') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Item2 = 'Act 2 Equip C';
        row.Prob2 = 5;
        row.Item3 = 'Act 2 Good';
        row.Prob3 = 1;
      }
      if (treasureClass === 'Andarielq (N)') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Item2 = 'Act 2 (N) Good';
        row.Prob2 = 1;
        row.Item3 = 'Act 1 (N) Equip C';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Andarielq (H)') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 1 (H) Equip C';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Baal') {
        SetQuestProp(row, 0);
        row.Item2 = 'Act 1 (N) Equip C';
        row.Item3 = 'Act 1 (N) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Baal (N)') {
        SetQuestProp(row, 0);
        row.Item1 = 'gld,mul=3536';
        row.Item2 = 'Act 1 (H) Equip C';
        row.Item3 = 'Act 1 (H) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Baal (H)') {
        SetQuestProp(row, 0);
        row.Item1 = 'gld,mul=4048';
        row.Item2 = 'Act 5 (H) Equip C';
        row.Item3 = 'Act 5 (H) Good';
        row.Prob3 = 3;
        row.Item4 = 'fed';
        row.Prob4 = 0;
        row.Item5 = '';
        row.Prob5 = '';
      }
      if (treasureClass === 'Baalq') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 25;
      }
      if (treasureClass === 'Baalq (N)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'Act 1 (H) Equip C';
        row.Prob1 = 26;
        row.Prob2 = 5;
      }
      if (treasureClass === 'Baalq (H)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'Act 5 (H) Equip C';
        row.Prob1 = 26;
        row.Prob2 = 5;
        row.Prob3 = 0;
        row.Item4 = 'uar'
        row.Prob4 = 1;
      }
      if (treasureClass === 'Blood Raven') {
        SetQuestProp(row, 6);
        row.Item2 = 'Act 1 Equip A';
        row.Prob2 = 14;
        row.Item3 = 'Act 1 Good';
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Blood Raven (N)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=3536';
        row.Prob1 = 6;
        row.Item2 = 'Act 1 (N) Equip A';
        row.Prob2 = 14;
        row.Item3 = 'Act 1 (N) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Blood Raven (H)') {
        row.Unique = 1024;
        row.Set = 850;
        row.Rare = 850;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Item1 = 'gld,mul=4048';
        row.Prob1 = 11;
        row.Item2 = 'Act 1 (H) Equip A';
        row.Prob2 = 33;
        row.Item3 = '6lw';
        row.Prob3 = 1;
        row.Prob4 = 9;
      }
      if (treasureClass === 'Countess') {
        row.Picks = 5;
        row.Unique = 1024;
        row.Magic = 800;
        row.Prob1 = 0;
        row.Prob2 = 5;
        row.Item3 = 'Act 2 Good';
        row.Prob3 = 1;
      }
      if (treasureClass === 'Countess (N)') {
        row.Picks = 5;
        row.Unique = 1024;
        row.Magic = 800;
        row.Prob1 = 0;
        row.Prob2 = 5;
        row.Item3 = 'Act 2 (N) Good';
        row.Prob3 = 1;
      }
      if (treasureClass === 'Countess (H)') {
        row.Picks = 5;
        row.Unique = 1024;
        row.Magic = 800;
        row.Prob1 = 0;
        row.Prob2 = 5;
        row.Item3 = 'Act 2 (H) Good';
        row.Prob3 = 1;
        row.Item4 = 'pk1';
        row.Prob4 = 1;
      }
      if (treasureClass === 'Countess Rune') {
        row.Item1 = 'Runes 5';
      }
      if (treasureClass === 'Countess Rune (N)') {
        row.Item1 = 'Runes 11';
        row.Prob1 = 3;
        row.Item2 = 'Runes 8';
        row.Prob2 = 1;
      }
      if (treasureClass === 'Countess Rune (H)') {
        row.Item1 = 'Runes 11';
        row.Prob1 = 6;
        row.Item2 = 'Runes 15';
        row.Prob2 = 2;
        row.Item3 = 'Runes 16';
        row.Prob3 = 1;
        row.Item4 = 'Runes 13';
        row.Prob4 = 6;
        row.Item5 = 'Runes 14';
        row.Prob5 = 2;
      }
      if (treasureClass === 'Council') {
        row.Unique = 999;
        row.Set = 997;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Item1 = 'gld,mul=3280';
        row.Prob1 = 6;
        row.Item3 = 'Act 4 Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
        row.Item5 = '';
        row.Prob5 = '';
      }
      if (treasureClass === 'Council (N)') {
        row.Unique = 999;
        row.Set = 997;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Item1 = 'gld,mul=4536';
        row.Prob1 = 6;
        row.Item3 = 'Act 4 (N) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
        row.Item5 = '';
        row.Prob5 = '';
      }
      if (treasureClass === 'Council (H)') {
        row.Unique = 999;
        row.Set = 997;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Item1 = 'gld,mul=5048';
        row.Prob1 = 6;
        row.Item3 = 'Act 4 (H) Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
        row.Item5 = '';
        row.Prob5 = '';
      }
      if (treasureClass === 'Cow') {
        row.Item3 = 'Act 5 Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Cow King') {
        row.Picks = 7;
      }
      if (treasureClass === 'Diablo') {
        SetQuestProp(row, 4);
        row.Prob2 = 15;
        row.Item3 = 'Act 5 Equip C';
        row.Prob3 = 3;
        row.Prob4 = 5;
      }
      if (treasureClass === 'Diabloq') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 5 Equip C';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Duriel') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Duriel (N)') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Duriel (H)') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Duriel - Base') {
        SetQuestProp(row, 11);
        row.Item3 = 'Act 3 Equip A';
        row.Prob3 = 3;
        row.Prob4 = 2;
      }
      if (treasureClass === 'Duriel (N) - Base') {
        SetQuestProp(row, 11);
        row.Item1 = 'gld,mul=3536';
        row.Item3 = 'Act 3 (N) Equip A';
        row.Prob3 = 3;
        row.Prob4 = 2;
      }
      if (treasureClass === 'Duriel (H) - Base') {
        SetQuestProp(row, 11);
        row.Item1 = 'gld,mul=4048';
        row.Item3 = 'Act 3 (H) Equip A';
        row.Prob3 = 3;
        row.Prob4 = 2;
        row.Prob5 = 0;
      }
      if (treasureClass === 'Durielq') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Durielq (N)') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Durielq (H)') {
        row.Picks = 7;
        row.Unique = 1024;
        row.Prob1 = 0;
      }
      if (treasureClass === 'Durielq - Base') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 3 Equip B';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Durielq (N) - Base') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 3 (N) Equip B';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Durielq (H) - Base') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 44;
        row.Prob2 = 3;
        row.Item3 = 'Act 3 (H) Equip B';
        row.Prob3 = 6;
        row.Item4 = 'xrs';
        row.Prob4 = 1;
      }
      if (treasureClass === 'Flying Scimitar') {
        row.Unique = 999;
        row.Set = 899;
        row.Rare = 850;
        row.Magic = 800;
        row.NoDrop = 0;
        row.Prob1 = 11;
        row.Prob2 = 7;
      }
      if (treasureClass === 'Griswold') {
        row.Picks = 4;
        SetDefaultQuestProp(row, false);
        row.Item1 = 'Act 2 Uitem C';
        row.Prob1 = 8;
        row.Item2 = 'Act 2 Melee A';
        row.Prob2 = 15;
        row.Item3 = 'bsd';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Griswold (N)') {
        row.Picks = 4;
        row.Unique = 999;
        row.Set = 999;
        row.Rare = 800;
        row.Magic = 800;
        row.Item1 = 'Act 1 (H) Uitem C';
        row.Prob1 = 8;
        row.Item2 = 'Act 1 (N) Melee B';
        row.Prob2 = 15;
      }
      if (treasureClass === 'Haphesto') {
        row.Picks = 4;
        SetQuestProp(row, 5);
        row.Item2 = 'Act 4 Equip A';
        row.Item3 = 'Act 5 Good';
        row.Prob3 = 4;
        row.Prob4 = 2;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Izual') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=2280';
        row.Prob1 = 8;
        row.Item3 = 'Act 4 Good';
        row.Prob3 = 8;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Mephisto') {
        SetQuestProp(row, 4);
        row.Prob2 = 15;
        row.Item3 = 'Act 5 Equip A';
        row.Prob3 = 3;
        row.Prob4 = 5;
      }
      if (treasureClass === 'Mephistoq') {
        SetDefaultQuestProp(row, true);
        row.Prob1 = 22;
        row.Prob2 = 1;
        row.Item3 = 'Act 5 Equip A';
        row.Prob3 = 3;
      }
      if (treasureClass === 'Nihlathak') {
        row.Picks = 6;
        SetQuestProp(row, 5);
        row.Item2 = 'Act 5 Equip C';
        row.Item3 = 'Act 5 Good';
        row.Prob3 = 3;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Radament') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=2280';
        row.Prob1 = 3;
        row.Item2 = 'Act 3 Equip A';
        row.Prob2 = 15;
        row.Item3 = 'Act 3 Good';
        row.Prob3 = 7;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Radament (N)') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=3536';
        row.Prob1 = 3;
        row.Item2 = 'Act 3 (N) Equip A';
        row.Prob2 = 3;
        row.Item3 = 'Act 3 (N) Good';
        row.Prob3 = 15;
        row.Item4 = '';
        row.Prob4 = '';
      }
      if (treasureClass === 'Smith') {
        row.Picks = 3;
        row.Unique = 1024;
        row.Set = 800;
        row.Rare = 800;
        row.Magic = 800;
      }
      if (treasureClass === 'Smith (N)') {
        row.Picks = 3;
        row.Unique = 1024;
        row.Set = 800;
        row.Rare = 800;
        row.Magic = 800;
      }
      if (treasureClass === 'Smith (H)') {
        row.Picks = 3;
        row.Unique = 1024;
        row.Set = 800;
        row.Rare = 800;
        row.Magic = 800;
      }
      if (treasureClass === 'Summoner') {
        SetDefaultQuestProp(row, true);
        row.Item1 = 'gld,mul=2280';
        row.Prob1 = 4;
        row.Item3 = 'Act 2 Good';
        row.Prob3 = 2;
        row.Item4 = 'Act 3 Equip A';
        row.Prob4 = 4;
        row.Item5 = '';
        row.Prob5 = '';
      }
    }

    // Increased more drop rates.
    {
      const chestLevel = ["A", "B", "C"];
      const diffLevel = ['', '(N) ', '(H) '];
      for (let acts = 1; acts <= 5; acts = acts + 1) {
        for (let level = 0; level < 3; level = level + 1) {
          for (let diff = 0; diff < 3; diff = diff + 1) {
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`) {
              if (acts === 5) {
                if (diff === 2) {
                  if (chestLevel[level] === "B") {
                    row.level = 70;
                  }
                  if (chestLevel[level] === "C") {
                    row.level = 70;
                    row.Prob1 = 0;
                    row.Prob2 = 0;
                    row.Prob3 = 0;
                    row.Prob4 = 0;
                    row.Prob6 = 5;
                    row.Prob7 = 5;
                    row.Prob8 = 5;
                  }
                }
              }
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Cast ${chestLevel[level]}`) {
              // Act 1
              if (acts === 1) {
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act 1 ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item4 = `Act 1 ${diffLevel[diff]}Magic B`;
                  }
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
                // Hell
                else if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act 1 ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item4 = `Act 1 ${diffLevel[diff]}Magic B`;
                  }
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
                // Normal
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic ${chestLevel[level]}`;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
              }
              // Act 2
              else if (acts === 2) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
                else if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
                else {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act 3 ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic C`;
                  }
                }
              }
              // Act 3
              else if (acts === 3) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic C`;
                  }
                }
              }
              else if (acts == 4) {
                if (chestLevel[level] === "A") {
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                }
                else {
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                }
              }
              else {
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic C`;
                  }
                }
                else {
                  if (chestLevel[level] === "A") {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic A`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Magic B`;
                  }
                }
                row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
              }
              row.Prob3 = 2;
              row.Prob4 = 7;
              row.Item5 = '';
              row.Prob5 = '';
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Chest ${chestLevel[level]}`) {
              // Act 1
              if (acts === 1) {
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 1 ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  }
                  else {
                    row.Item2 = `Act 1 ${diffLevel[diff]}Equip B`;
                  }
                  row.Prob2 = 10;
                }
                // Hell
                else if (diff === 2) {
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 1 ${diffLevel[diff]}Equip B`;
                  }
                  else {
                    row.Item2 = `Act 1 ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  }
                  row.Prob2 = 10;
                }
                // Normal
                else {
                  if (chestLevel[level] === "C") {
                    row.Prob2 = 15;
                  }
                  else {
                    row.Prob2 = 10;
                  }
                  row.Item2 = `Act 1 ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                }
              }
              // Act 2
              else if (acts === 2) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 2 (H) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 2 (H) Equip B`;
                  }
                  row.Prob2 = 10;
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 2 (N) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 2 (N) Equip B`;
                  }
                  row.Prob2 = 10;
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 2 Equip A`;
                    row.Prob2 = 10;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item2 = `Act 2 Equip B`;
                    row.Prob2 = 10;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 2 Equip C`;
                    row.Prob2 = 15;
                  }
                }
              }
              // Act 3
              else if (acts === 3) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 3 (H) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 3 (H) Equip B`;
                  }
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 3 (N) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 3 (N) Equip B`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 3 Equip A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item2 = `Act 3 Equip B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 3 Equip C`;
                  }
                }
                row.Prob2 = 10;
              }
              // Act 4
              else if (acts === 4) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 4 (H) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 4 (H) Equip B`;
                  }
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 4 (N) Equip A`;
                  }
                  if (chestLevel[level] === "B") {
                    row.Item2 = `Act 4 (N) Equip B`;
                  }
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 4 (N) Equip C`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 4 Equip A`;
                  }
                  else {
                    row.Item2 = `Act 4 Equip B`;
                  }
                }
                row.Prob2 = 10;
              }
              // Act
              else if (acts === 5) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 5 (H) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 5 (H) Equip B`;
                  }
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 5 (N) Equip A`;
                  }
                  else {
                    row.Item2 = `Act 5 (N) Equip B`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "A") {
                    row.Item2 = `Act 5 Equip A`;
                  }
                  else {
                    row.Item2 = `Act 5 Equip B`;
                  }
                }
                row.Prob2 = 10;
              }
              else {
                row.Prob2 = 15;
              }
              row.NoDrop = 50;
              row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
              row.Prob3 = 4;
              row.Item4 = '';
              row.Prob4 = '';
              row.Item5 = '';
              row.Prob5 = '';
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`) {
              if (acts === 5) {
                if (diff === 2) {
                  if (chestLevel[level] === "B") {
                    row.Prob1 = 1;
                    row.Prob3 = 4;
                    row.Prob4 = 6;
                    row.Prob5 = 12;
                    row.Prob6 = 6;
                    row.Prob7 = 4;
                    row.Prob8 = 3;
                  }
                  if (chestLevel[level] === "C") {
                    row.Prob1 = 1;
                    row.Prob3 = 4;
                    row.Prob4 = 6;
                    row.Prob5 = 11;
                    row.Prob6 = 5;
                    row.Prob7 = 5;
                    row.Prob8 = 5;
                  }
                }
              }
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}H2H ${chestLevel[level]}`) {
              if (acts === 2) {
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act 3 Good`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                }
                else {
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
              }
              else {
                row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
              }
              row.Prob3 = 2;
              row.Item4 = '';
              row.Prob4 = '';
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Melee ${chestLevel[level]}`) {
              if (acts === 5) {
                if (diff === 2) {
                  if (chestLevel[level] === "B") {
                    row.Prob1 = 2;
                    row.Prob3 = 6;
                    row.Prob4 = 3;
                    row.Prob5 = 14;
                    row.Prob6 = 7;
                    row.Prob7 = 3;
                    row.Prob8 = 3;
                  }
                  if (chestLevel[level] === "C") {
                    row.Prob1 = 1;
                    row.Prob3 = 3;
                    row.Prob4 = 3;
                    row.Prob5 = 11;
                    row.Prob6 = 7;
                    row.Prob7 = 4;
                    row.Prob8 = 4;
                  }
                }
              }
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Miss ${chestLevel[level]}`) {
              row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
              row.Prob3 = 2;
              if (acts === 1) {
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                // Hell
                else if (diff === 2) {
                  if (chestLevel[level] === "A") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow A`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                }
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                }
              }
              else if (acts === 2) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act 3 ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
              }
              else if (acts === 3) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item3 = `Act 3 ${diffLevel[diff]}Good`;
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
              }
              else if (acts === 5) {
                // Hell
                if (diff === 2) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                else {
                  row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                }
                // Nightmare
                if (diff === 1) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
                // Normal
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow B`;
                  }
                  else {
                    row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
                  }
                }
              }
              else {
                row.Item4 = `Act ${acts} ${diffLevel[diff]}Bow ${chestLevel[level]}`;
              }
              row.Prob4 = 6;
              row.Item5 = 'Ammo';
              row.Prob5 = 3;
              row.Item6 = '';
              row.Prob6 = '';
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super ${chestLevel[level]}`) {
              // Super A
              if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super A`) {
                // Act 1
                if (acts === 1) {
                  // Normal
                  if (diff === 0) {
                    row.level = 5;
                    row.Prob1 = 14;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Prob1 = 14;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Prob1 = 12;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 6;
                }
                // Act 2
                if (acts === 2) {
                  // Normal
                  if (diff === 0) {
                    row.level = 12;
                    row.Item2 = `Act 1 Equip C`;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Item2 = `Act 1 (N) Equip C`;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Item2 = `Act 1 (H) Equip C`;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Prob2 = 7;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 6;
                }
                // Act 3
                if (acts === 3) {
                  // Normal
                  if (diff === 0) {
                    row.Item2 = `Act 2 Equip C`;
                    row.Prob1 = 15;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Item2 = `Act 2 (N) Equip C`;
                    row.Prob1 = 15;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Item2 = `Act 2 (H) Equip C`;
                    row.Prob1 = 15;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob2 = 7;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 6;
                }
                // Act 4
                if (acts === 4) {
                  // Normal
                  if (diff === 0) {
                    row.level = 26;
                    row.Prob1 = 12;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Prob1 = 12;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Prob1 = 12;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  row.Prob2 = 12;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 3;
                }
                // Act 5
                if (acts === 5) {
                  // Hell
                  if (diff === 2) {
                    row.Item2 = `Act 4 ${diffLevel[diff]}Equip C`;
                  }
                  else if (diff === 0) {
                    row.Item2 = `Act 4 ${diffLevel[diff]}Equip C`;
                  }
                  else {
                    row.Item2 = `Act 4 ${diffLevel[diff]}Equip C`;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Prob2 = 7;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 3;
                }
              }
              // Super B
              else if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super B`) {
                if (acts === 1) {
                  // Normal
                  if (diff === 0) {
                    row.level = 7;
                    row.Prob1 = 14;
                    row.Prob3 = 6;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Prob1 = 14;
                    row.Prob3 = 6;
                  }
                  // Hell
                  if (diff === 2) {
                    row.Prob1 = 12;
                    row.Prob3 = 6;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                }
                if (acts === 2) {
                  row.Prob1 = 15;
                  row.Prob3 = 6;
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                }
                if (acts === 3) {
                  // Hell
                  if (diff === 2) {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  }
                  else if (diff === 1) {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  }
                  else {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  }
                  row.Prob1 = 15;
                  row.Prob3 = 6;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                }
                if (acts === 4) {
                  if (diff === 1) {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                    row.Prob1 = 12;
                    row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                    row.Prob3 = 3;
                  }
                  else if (diff === 2) {
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                    row.Prob1 = 12;
                    row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                    row.Prob3 = 3;
                  }
                  else {
                    row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                    row.Prob1 = 12;
                    row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                    row.Prob3 = 3;
                  }
                  row.Prob2 = 12;
                }
                if (acts === 5) {
                  row.Prob1 = 15;
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip C`;
                  row.Prob3 = 3;
                }
                row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;

              }
              // Super C
              else if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super C`) {
                // Act 1
                if (acts === 1) {
                  // Hell
                  if (diff === 2) {
                    row.Prob2 = 2;
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  else {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Item2 = `Act 2 ${diffLevel[diff]}Equip A`;
                  row.Prob3 = 6;
                }
                // Act 2
                if (acts === 2) {
                  // Hell
                  if (diff === 2) {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  // Nightmare
                  if (diff === 1) {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  // Normal
                  if (diff === 0) {
                    row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Item2 = `Act 3 ${diffLevel[diff]}Equip A`;
                  row.Prob3 = 6;
                } // Act 3
                if (acts === 3) {
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip A`;
                  row.Prob1 = 15;
                  row.Item2 = `Act 4 ${diffLevel[diff]}Equip A`;
                  row.Prob2 = 2;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 6;
                }
                // Act 4
                if (acts === 4) {
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 12;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Equip B`;
                  row.Prob2 = 12;
                  row.Item3 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob3 = 3;
                }
                // Act 5
                if (acts === 5) {
                  row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
                  row.Prob1 = 15;
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Good`;
                  row.Prob2 = 3;
                }
              }
              row.Picks = 5;
              row.Set = 1024;
              row.Rare = 800;
              row.Magic = 800;
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Super ${chestLevel[level]}x`) {
              row.group = 18;
              row.Picks = 5;
              row.Magic = 900;
            }
            if (treasureClass === `Act ${acts} ${diffLevel[diff]}Unique ${chestLevel[level]}`) {
              if (acts === 2) {
                if (diff === 0) {
                  if (chestLevel[level] === "C") {
                    row.Item2 = `Act 3 ${diffLevel[diff]}Good`;
                  }
                  else {
                    row.Item2 = `Act ${acts} ${diffLevel[diff]}Good`;
                  }
                }
                else {
                  row.Item2 = `Act ${acts} ${diffLevel[diff]}Good`;
                }
              }
              else {
                row.Item2 = `Act ${acts} ${diffLevel[diff]}Good`;
              }
              row.Picks = 3;
              row.Unique = 933;
              row.Set = 1010;
              row.Rare = 800;
              row.Magic = 800;
              row.Item1 = `Act ${acts} ${diffLevel[diff]}Equip ${chestLevel[level]}`;
              row.Prob1 = 11;
              row.Prob2 = 1;
            }
          }
        }
      }
    }

    if (treasureClass === 'Chipped Gem') {
      row.Prob1 = 10;
      row.Prob2 = 6;
      row.Prob3 = 6;
      row.Prob4 = 6;
      row.Prob5 = 6;
      row.Prob6 = 6;
      row.Prob7 = 6;
    }
    if (treasureClass === 'Flawed Gem') {
      row.Prob7 = 3;
    }
    if (treasureClass === 'Normal Gem') {
      row.Prob7 = 3;
    }
    if (treasureClass === 'Flawless Gem') {
      row.Prob7 = 3;
    }
    if (treasureClass === 'Perfect Gem') {
      row.Prob7 = 3;
    }
    if (treasureClass === 'Runes 12') {
      row.Item4 = 'Runes 2';
    }
    if (treasureClass === 'Runes 13') {
      row.Item4 = 'Runes 2';
    }
    if (treasureClass === 'Runes 14') {
      row.Item4 = 'Runes 2';
    }
    if (treasureClass === 'Runes 15') {
      row.Item4 = 'Runes 2';
    }
    if (treasureClass === 'Runes 16') {
      row.Item4 = 'Runes 2';
    }
  });
  D2RMM.writeTsv(treasureclassexFilename, treasureclassex);
}

// ShowItemLevel
{
  if (config.weapons) {
    const miscFilename = 'global\\excel\\weapons.txt';
    const misc = D2RMM.readTsv(miscFilename);
    misc.rows.forEach((row) => {
      if (
        // don't modify throwing potions (gas, oil pots)
        row.type !== 'tpot'
      ) {
        row.ShowLevel = 1;
      }
    });
    D2RMM.writeTsv(miscFilename, misc);
  }

  if (config.armor) {
    const miscFilename = 'global\\excel\\armor.txt';
    const misc = D2RMM.readTsv(miscFilename);
    misc.rows.forEach((row) => {
      row.ShowLevel = 1;
    });
    D2RMM.writeTsv(miscFilename, misc);
  }

  if (config.jewelry || config.charms || config.jewels) {
    const miscFilename = 'global\\excel\\misc.txt';
    const misc = D2RMM.readTsv(miscFilename);
    misc.rows.forEach((row) => {
      if (config.jewelry && ['amu', 'rin'].indexOf(row.code) !== -1) {
        row.ShowLevel = 1;
      }
      if (config.charms && ['cm1', 'cm2', 'cm3'].indexOf(row.code) !== -1) {
        row.ShowLevel = 1;
      }
      if (config.jewels && ['jew'].indexOf(row.code) !== -1) {
        row.ShowLevel = 1;
      }
    });
    D2RMM.writeTsv(miscFilename, misc);
  }

}

// Removed the gems from the rune upgrade recipes. 
// - Up until Pul: 3 of the same runes = next rune
// - After Pul: 2 of a kind + jewel = next rune
// Removed Hel-Rune from the de-socked recipe. New recipe: 1 Town Scroll + Socketed Item (destroys gems).
{
  const cubemainFilename = 'global\\excel\\cubemain.txt';
  const cubemain = D2RMM.readTsv(cubemainFilename);
  cubemain.rows.forEach((row) => {
    if (row.description === '3 Thul Runes + 1 Chipped Topaz -> Amn Rune') {
      row.description = '3 Thul Runes -> Amn Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Amn Runes + 1 Chipped Amethyst -> Sol Rune') {
      row.description = '3 Amn Runes -> Sol Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Sol Runes + 1 Chipped Sapphire -> Shael Rune') {
      row.description = '3 Sol Runes -> Shael Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Shael Runes + 1 Chipped Ruby -> Dol Rune') {
      row.description = '3 Shael Runes -> Dol Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Dol Runes + 1 Chipped Emerald -> Hel Rune') {
      row.description = '3 Dol Runes -> Hel Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Hel Runes + 1 Chipped Diamond -> Io Rune') {
      row.description = '3 Hel Runes -> Io Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Io Runes + 1 Flawed Topaz -> Lum Rune') {
      row.description = '3 Io Runes -> Lum Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Lum Runes + 1 Flawed Amethyst -> Ko Rune') {
      row.description = '3 Lum Runes -> Ko Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Ko Runes + 1 Flawed Sapphire -> Fal Rune') {
      row.description = '3 Ko Runes -> Fal Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Fal Runes + 1 Flawed Ruby -> Lem Rune') {
      row.description = '3 Fal Runes -> Lem Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '3 Lem Runes + 1 Flawed Emerald -> Pul Rune') {
      row.description = '3 Lem Runes -> Pul Rune';
      row.numinputs = 3;
      row['input 2'] = '';
    }
    if (row.description === '2 Pul Runes + 1 Flawed Diamond -> Um Rune') {
      row.description = '2 Pul Runes + 1 Jewel -> Um Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Um Runes + 1 Standard Topaz -> Mal Rune') {
      row.description = '2 Um Runes + 1 Jewel -> Mal Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Mal Runes + 1 Standard Amethyst -> Ist Rune') {
      row.description = '2 Mal Runes + 1 Jewel -> Ist Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Ist Runes + 1 Standard Sapphire -> Gul Rune') {
      row.description = '2 Ist Runes + 1 Jewel -> Gul Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Gul Runes + 1 Standard Ruby -> Vex Rune') {
      row.description = '2 Gul Runes + 1 Jewel -> Vex Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Vex Runes + 1 Standard Emerald -> Ohm Rune') {
      row.description = '2 Vex Runes + 1 Jewel -> Ohm Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Ohm Runes + 1 Standard Diamond -> Lo Rune') {
      row.description = '2 Ohm Runes + 1 Jewel -> Lo Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Lo Runes + 1 Flawless Topaz -> Sur Rune') {
      row.description = '2 Lo Runes + 1 Jewel -> Sur Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Sur Runes + 1 Flawless Amethyst -> Ber Rune') {
      row.description = '2 Sur Runes + 1 Jewel -> Ber Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Ber Runes + 1 Flawless Sapphire -> Jah Rune') {
      row.description = '2 Ber Runes + 1 Jewel -> Jah Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Jah Runes + 1 Flawless Ruby -> Cham Rune') {
      row.description = '2 Jah Runes + 1 Jewel -> Cham Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '2 Cham Runes + 1 Flawless Emerald -> Zod Rune') {
      row.description = '2 Cham Runes + Jewel -> Zod Rune';
      row['input 2'] = 'jew';
    }
    if (row.description === '1 Hel Rune + Scroll of Town Portal + 1 Socketed Item -> Clear Sockets on Item') {
      row.description = 'Scroll of Town Portal + 1 Socketed Item -> Clear Sockets on Item';
      row.numinputs = 2;
      row['input 1'] = 'any,sock';
      row['input 2'] = 'tsc';
      row['input 3'] = '';
    }
    if (row.description === '1 Twisted Essence of Suffering + 1 Charged Essence of Hatred + 1 Burning Essence of Terror + 1 Festering Essence of Destruction -> Token of Absolution') {
      row.description = 'Token of Absolution',
        row.numinputs = 2;
      row['input 1'] = 'tsc';
      row['input 2'] = 'isc';
      row['input 3'] = '';
      row['input 4'] = '';
    }
  });
  D2RMM.writeTsv(cubemainFilename, cubemain);
}

// Socket recipes
{
  const cubemainFilename = 'global\\excel\\cubemain.txt';
  const cubemain = D2RMM.readTsv(cubemainFilename);
  // Add upto 4 Sockets with number of Perfect Gems.  5 & 6 Sockets use Perfect Skulls
  for (let sockets = 1; sockets <= 6; sockets = sockets + 1) {
    const socketRecipe = {
      description: `${sockets} Perfect Gem`,
      enabled: 1,
      version: 100,
      numinputs: sockets + 1,
      'input 2': `gem4,qty=${sockets}`,
      output: 'useitem',
      'mod 1': 'sock',
      'mod 1 min': sockets,
      'mod 1 max': sockets,
      '*eol': 0,
    };
    function addRecipeNormal(code, name) {
      cubemain.rows.push({
        ...socketRecipe,
        description: `${socketRecipe.description} + Normal ${name} -> Socket Normal ${name} ${sockets}`,
        'input 1': `${code},nor,nos`
      });
    }
    function addRecipeSpecial(code, name) {
      cubemain.rows.push({
        ...socketRecipe,
        description: `${socketRecipe.description} + Special ${name} -> Socket Special ${name} ${sockets}`,
        'input 1': `${code},hiq,nos`
      });
    }
    function addRecipeNormalSkull(code, quantity) {
      cubemain.rows.push({
        ...socketRecipe,
        description: `Perfect Skull + Normal or Unique Weapon -> Socket Normal or Unique ${sockets}`,
        numinputs: quantity + 2,
        'input 1': `weap,${code},nos`,
        'input 2': `ring,uni`,
        'input 3': `skz,qty=${quantity}`,
      });
    }
    if (sockets <= 4) {
      addRecipeNormal('helm', 'Helm');
      addRecipeNormal('shld', 'Shield');
      addRecipeNormal('tors', 'Torso');
      addRecipeNormal('weap', 'Weapon');
      addRecipeSpecial('helm', 'Helm');
      addRecipeSpecial('shld', 'Shield');
      addRecipeSpecial('tors', 'Torso');
      addRecipeSpecial('weap', 'Weapon');
    }
    if (sockets === 5) {
      addRecipeNormalSkull('nor', 1);
      addRecipeNormalSkull('hiq', 1);
    }
    if (sockets === 6) {
      addRecipeNormalSkull('nor', 2);
      addRecipeNormalSkull('hiq', 2);
    }
  }

  // Perfect Skull and Ring
  {
    const socketRecipe = {
      description: `1 Unique Rings`,
      enabled: 1,
      version: 100,
      numinputs: 3,
      'input 2': 'ring,uni',
      'input 3': 'skz,qty=1',
      'mod 1': 'sock',
      'mod 1 min': 5,
      'mod 1 max': 5,
      output: 'useitem',
      '*eol': 0,
    };

    function addRecipeRingPerfectSkullNormal(code, name, nRings) {
      const sockets = nRings === 1 ? 5 : 6;
      cubemain.rows.push({
        ...socketRecipe,
        description: `${socketRecipe.description} ${nRings} Perfect Skull + Normal ${name} -> Socket ${name} ${sockets}`,
        'input 1': `${code},nor,nos`,
        numinputs: nRings + 2,
        'input 3': `skz,qty=${nRings}`,
        'mod 1 min': sockets,
        'mod 1 max': sockets,
      });
    }

    function addRecipeRingPerfectSkullSpecial(code, name, nRings) {
      const sockets = nRings === 1 ? 5 : 6;
      cubemain.rows.push({
        ...socketRecipe,
        description: `${socketRecipe.description} ${nRings} Perfect Skull + Special ${name} -> Socket ${name} ${sockets}`,
        'input 1': `${code},hiq,nos`,
        numinputs: nRings + 2,
        'input 3': `skz,qty=${nRings}`,
        'mod 1 min': sockets,
        'mod 1 max': sockets,
      });
    }

    addRecipeRingPerfectSkullNormal('weap', 'Weapon', 1);
    addRecipeRingPerfectSkullNormal('weap', 'Weapon', 2);
    addRecipeRingPerfectSkullSpecial('weap', 'Weapon', 1);
    addRecipeRingPerfectSkullSpecial('weap', 'Weapon', 2);
  }

  // Perfect Skull and Set Item
  {
    const socketRecipe = {
      description: `1 Perfect Skull + 1 Set Item -> 1 Socket Set Item`,
      enabled: 1,
      version: 100,
      numinputs: 2,
      'input 1': `any,set`,
      'input 2': 'skz,qty=1',
      'mod 1': `sock`,
      'mod 1 min': 1,
      'mod 1 max': 1,
      output: 'useitem',
      '*eol': 0,
    };

    function addRecipe() {
      cubemain.rows.push({ ...socketRecipe });
    }

    addRecipe();
  }

  // Perfect Gem and Unique Item
  {
    const socketRecipe = {
      description: `1 Perfect Gem + 1 Unique Item -> 1 Socket Set Item`,
      enabled: 1,
      version: 100,
      numinputs: 2,
      'input 1': `any,uni`,
      'input 2': 'gem4,qty=1',
      'mod 1': `sock`,
      'mod 1 min': 1,
      'mod 1 max': 1,
      output: 'useitem',
      '*eol': 0,
    };

    function addRecipe() {
      cubemain.rows.push({ ...socketRecipe });
    }

    addRecipe();
  }
  D2RMM.writeTsv(cubemainFilename, cubemain);
}

// MrLamaSc character startup
{
  const charstatsFilename = 'global\\excel\\charstats.txt';
  const charstats = D2RMM.readTsv(charstatsFilename);
  charstats.rows.forEach((row) => {
    const id = row.class;
    if (id === 'Sorceress') {
      row.StartSkill = 'charged bolt';
    }
    if (id === 'Amazon' || id === 'Assassin' || id === 'Barbarian' || id === 'Druid' || id === 'Paladin') {
      row.item2 = 'hp1';
      row.item2loc = '';
      row.item2count = 4;
      row.item3 = 'tsc';
      row.item3count = 1;
      row.item4 = 'isc';
      row.item5 = 0;
      row.item5count = 0;
    }
  });
  D2RMM.writeTsv(charstatsFilename, charstats);
}

// Move Cain
{
  if (config["act5_options"] == "stash") {
    D2RMM.copyFile(
      'act5_cain_stash\\global\\tiles\\expansion\\combined_ds1.bin',
      'global\\tiles\\expansion\\combined_ds1.bin',
      true
    );

    D2RMM.copyFile(
      'act5_cain_stash\\global\\tiles\\expansion\\town\\townwest.ds1',
      'global\\tiles\\expansion\\town\\townwest.ds1',
      true
    );
  }

  if (config["act5_options"] == "malah") {
    D2RMM.copyFile(
      'act5_cain_malah\\global\\tiles\\expansion\\combined_ds1.bin',
      'global\\tiles\\expansion\\combined_ds1.bin',
      true
    );

    D2RMM.copyFile(
      'act5_cain_malah\\global\\tiles\\expansion\\town\\townwest.ds1',
      'global\\tiles\\expansion\\town\\townwest.ds1',
      true
    );
  }

  if (config["act5_options"] == "larzuk") {
    D2RMM.copyFile(
      'act5_cain_larzuk\\global\\tiles\\expansion\\combined_ds1.bin',
      'global\\tiles\\expansion\\combined_ds1.bin',
      true
    );

    D2RMM.copyFile(
      'act5_cain_larzuk\\global\\tiles\\expansion\\town\\townwest.ds1',
      'global\\tiles\\expansion\\town\\townwest.ds1',
      true
    );
  }
}

// Copy files from hd to hd
{
  D2RMM.copyFile(
    'hd', // <mod folder>\hd
    'hd', // <diablo 2 folder>\mods\<modname>\<modname>.mpq\data\hd
    true // overwrite any conflicts
  );
}

// Modify Stash
{
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
}
