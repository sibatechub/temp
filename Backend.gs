/*************************************************
 * DRIEMS IT DASHBOARD
 * COMPUTER MANAGEMENT MODULE
 *
 * FILE: ComputerManagement.gs
 *************************************************/


/*************************************************
 * CONFIGURATION
 *************************************************/

// Separate Spreadsheet containing Computer Details
const COMPUTER_SPREADSHEET_ID =
  "1w7UhKvXK_zk-M3qXG0cCrc_FVdW3eOQ4Pei7lwiSy5A";


// Computer Details Sheet Name
const COMPUTER_MASTER_SHEET = "Master";


// Separate Spreadsheet containing:
// User Details / Authentication / Building Assignment
const USER_DETAILS_SPREADSHEET_ID =
  "1_Cu4-Z35C0FXpBJsApxmNpwimypnPGij7rXnS8c8524";


// Building Assignment Sheet Name
const BUILDING_TECHNICIAN_MAP_SHEET =
  "Building_Technician_Map";


/*************************************************
 * GET COMPUTER SPREADSHEET
 *************************************************/
function getComputerSpreadsheet() {

  return SpreadsheetApp.openById(
    COMPUTER_SPREADSHEET_ID
  );

}


/*************************************************
 * GET COMPUTER MASTER SHEET
 *************************************************/
function getComputerMasterSheet() {

  const ss = getComputerSpreadsheet();

  return ss.getSheetByName(
    COMPUTER_MASTER_SHEET
  );

}


/*************************************************
 * GET BUILDING TECHNICIAN MAP SHEET
 *
 * This sheet is located in the existing
 * Authentication / IT Dashboard Spreadsheet.
 *************************************************/
function getBuildingTechnicianMapSheet() {

  return SpreadsheetApp
    .openById(
      USER_DETAILS_SPREADSHEET_ID
    )
    .getSheetByName(
      BUILDING_TECHNICIAN_MAP_SHEET
    );

}


/*************************************************
 * GET CURRENT USER SESSION
 *************************************************/
function getComputerUserSession(sessionToken) {

  if (!sessionToken) {
    throw new Error("Session token is required");
  }

  const session =
    getUserSession(sessionToken);

  if (!session) {
    throw new Error("Session expired");
  }

  return session;

}


/*************************************************
 * GET ASSIGNED BUILDINGS
 *
 * ADMIN:
 *   No restriction
 *
 * TECHNICIAN:
 *   Gets only assigned buildings
 *************************************************/
function getAssignedBuildings(sessionToken) {

  const session =
    getComputerUserSession(
      sessionToken
    );


  const role =
    String(session.role || "")
      .trim()
      .toUpperCase();


  /***********************************************
   * ADMIN
   ***********************************************/
  if (role === "ADMIN") {

    return {
      isAdmin: true,
      buildings: []
    };

  }


  /***********************************************
   * GET BUILDING ASSIGNMENT DATA
   ***********************************************/
  const sheet =
    getBuildingTechnicianMapSheet();


  if (!sheet) {
    throw new Error(
      "Building_Technician_Map sheet not found"
    );
  }


  const data =
    sheet.getDataRange().getValues();


  const technicianEmail =
    String(session.email || "")
      .trim()
      .toLowerCase();


  const buildings = [];


  /***********************************************
   * BUILDING MAP COLUMNS
   *
   * 0 = Sl No
   * 1 = Building Name
   * 2 = Technician Name
   * 3 = Technician Email
   * 4 = Mobile No
   * 5 = Active Status
   ***********************************************/
  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const buildingName =
      String(data[i][1] || "")
        .trim();


    const assignedEmail =
      String(data[i][3] || "")
        .trim()
        .toLowerCase();


    const activeStatus =
      String(data[i][5] || "")
        .trim()
        .toUpperCase();


    // Technician must match
    // and building assignment must be active
    if (
      assignedEmail === technicianEmail &&
      activeStatus === "YES"
    ) {

      if (
        buildingName &&
        !buildings.includes(
          buildingName
        )
      ) {

        buildings.push(
          buildingName
        );

      }

    }

  }


  return {

    isAdmin: false,

    buildings: buildings

  };

}


/*************************************************
 * GET ALL COMPUTER DATA
 *
 * ADMIN:
 *   All computer records
 *
 * TECHNICIAN:
 *   Only assigned building computer records
 *************************************************/
function getFilteredComputerData(sessionToken) {

  /***********************************************
   * VALIDATE SESSION
   ***********************************************/
  const session =
    getComputerUserSession(
      sessionToken
    );


  const role =
    String(session.role || "")
      .trim()
      .toUpperCase();


  /***********************************************
   * GET COMPUTER MASTER DATA
   ***********************************************/
  const sheet =
    getComputerMasterSheet();


  if (!sheet) {
    throw new Error(
      "Computer Master sheet not found"
    );
  }


  const data =
    sheet.getDataRange().getValues();


  // Empty Sheet
  if (data.length <= 1) {
    return [];
  }


  // Remove Header
  const rows = data.slice(1);


  /***********************************************
   * ADMIN ACCESS
   ***********************************************/
  if (role === "ADMIN") {

    return rows;

  }


  /***********************************************
   * TECHNICIAN ACCESS
   ***********************************************/
  if (role === "TECHNICIAN") {

    const access =
      getAssignedBuildings(
        sessionToken
      );


    const assignedBuildings =
      access.buildings || [];


    // Technician has no assigned building
    if (
      assignedBuildings.length === 0
    ) {

      return [];

    }


    /***********************************************
     * COMPUTER MASTER
     *
     * COLUMN INDEX 1 = BUILDING NAME
     ***********************************************/
    return rows.filter(function(row) {

      const building =
        String(row[1] || "")
          .trim();


      return assignedBuildings.includes(
        building
      );

    });

  }


  /***********************************************
   * OTHER ROLES
   ***********************************************/
  return [];

}


/*************************************************
 * CONVERT COMPUTER ROW TO OBJECT
 *
 * COMPUTER MASTER COLUMN ORDER:
 *
 * 0  = SL NO
 * 1  = Building Name
 * 2  = Lab / Room Name
 * 3  = Computer No
 * 4  = CPU Brand
 * 5  = Monitor Brand
 * 6  = Keyboard Brand
 * 7  = Mouse Brand
 * 8  = Motherboard
 * 9  = Processor
 * 10 = Primary Storage
 * 11 = RAM
 * 12 = Graphic Card
 * 13 = MAC Address
 * 14 = OS
 * 15 = Level
 * 16 = QR Code
 * 17 = Direct Link
 * 18 = QR Code Link
 * 19 = Timestamp
 * 20 = User Name
 * 21 = Technician Email
 * 22 = Secondary Storage
 * 23 = External Storage
 * 24 = Floor
 * 25 = Asset Type
 * 26 = Printer Brand
 * 27 = Printer Model
 * 28 = Types of Printer
 *************************************************/
function computerRowToObject(row) {

  return {

    slNo: row[0] || "",

    buildingName: row[1] || "",

    labRoomName: row[2] || "",

    computerNo: row[3] || "",

    cpuBrand: row[4] || "",

    monitorBrand: row[5] || "",

    keyboardBrand: row[6] || "",

    mouseBrand: row[7] || "",

    motherboard: row[8] || "",

    processor: row[9] || "",

    primaryStorage: row[10] || "",

    ram: row[11] || "",

    graphicCard: row[12] || "",

    macAddress: row[13] || "",

    os: row[14] || "",

    level: row[15] || "",

    qrCode: row[16] || "",

    directLink: row[17] || "",

    qrCodeLink: row[18] || "",

    timestamp:
  row[19] instanceof Date
    ? Utilities.formatDate(
        row[19],
        Session.getScriptTimeZone(),
        "dd-MM-yyyy HH:mm:ss"
      )
    : String(row[19] || ""),

    userName: row[20] || "",

    technicianEmail: row[21] || "",

    secondaryStorage: row[22] || "",

    externalStorage: row[23] || "",

    floor: row[24] || "",

    assetType: row[25] || "",

    printerBrand: row[26] || "",

    printerModel: row[27] || "",

    printerType: row[28] || ""

  };

}


/*************************************************
 * GET COMPUTER LIST
 *
 * This is the main function that the frontend
 * will call to load computers.
 *************************************************/
function getComputerList(
  sessionToken,
  filter
) {

  filter = filter || {};


  const data =
    getFilteredComputerData(
      sessionToken
    );


  const search =
    String(filter.search || "")
      .trim()
      .toLowerCase();


  const building =
    String(filter.building || "")
      .trim();


  const labRoom =
    String(filter.labRoom || "")
      .trim();


  const floor =
    String(filter.floor || "")
      .trim();


  const assetType =
    String(filter.assetType || "")
      .trim();


  const os =
    String(filter.os || "")
      .trim();


  const cpuBrand =
    String(filter.cpuBrand || "")
      .trim();


  const processor =
    String(filter.processor || "")
      .trim();


  const ram =
    String(filter.ram || "")
      .trim();


  const primaryStorage =
    String(filter.primaryStorage || "")
      .trim();


  const graphicCard =
    String(filter.graphicCard || "")
      .trim();


  const userName =
    String(filter.userName || "")
      .trim();


  const filtered =
    data.filter(function(row) {


      /*********************************************
       * SEARCH
       *********************************************/
      if (search) {

        const searchableText = [

          row[3],   // Computer No
          row[13],  // MAC Address
          row[20],  // User Name
          row[9],   // Processor
          row[1],   // Building
          row[2]    // Lab / Room

        ]
          .join(" ")
          .toLowerCase();


        if (
          searchableText.indexOf(
            search
          ) === -1
        ) {

          return false;

        }

      }


      /*********************************************
       * BUILDING
       *********************************************/
      if (
        building &&
        String(row[1]).trim() !== building
      ) {

        return false;

      }


      /*********************************************
       * LAB / ROOM
       *********************************************/
      if (
        labRoom &&
        String(row[2]).trim() !== labRoom
      ) {

        return false;

      }


      /*********************************************
       * FLOOR
       *********************************************/
      if (
        floor &&
        String(row[24]).trim() !== floor
      ) {

        return false;

      }


      /*********************************************
       * ASSET TYPE
       *********************************************/
      if (
        assetType &&
        String(row[25]).trim() !== assetType
      ) {

        return false;

      }


      /*********************************************
       * OS
       *********************************************/
      if (
        os &&
        String(row[14]).trim() !== os
      ) {

        return false;

      }


      /*********************************************
       * CPU BRAND
       *********************************************/
      if (
        cpuBrand &&
        String(row[4]).trim() !== cpuBrand
      ) {

        return false;

      }


      /*********************************************
       * PROCESSOR
       *********************************************/
      if (
        processor &&
        String(row[9]).trim() !== processor
      ) {

        return false;

      }


      /*********************************************
       * RAM
       *********************************************/
      if (
        ram &&
        String(row[11]).trim() !== ram
      ) {

        return false;

      }


      /*********************************************
       * PRIMARY STORAGE
       *********************************************/
      if (
        primaryStorage &&
        String(row[10]).trim() !==
        primaryStorage
      ) {

        return false;

      }


      /*********************************************
       * GRAPHIC CARD
       *********************************************/
      if (
        graphicCard &&
        String(row[12]).trim() !==
        graphicCard
      ) {

        return false;

      }


      /*********************************************
       * USER NAME
       *********************************************/
      if (
        userName &&
        String(row[20]).trim() !==
        userName
      ) {

        return false;

      }


      return true;

    });


  return filtered.map(
    computerRowToObject
  );

}


/*************************************************
 * GET SINGLE COMPUTER DETAILS
 *
 * Search by Computer Number
 *************************************************/
function getComputerDetails(
  sessionToken,
  computerNo
) {

  if (!computerNo) {

    throw new Error(
      "Computer number is required"
    );

  }


  const data =
    getFilteredComputerData(
      sessionToken
    );


  const searchComputerNo =
    String(computerNo)
      .trim()
      .toLowerCase();


  for (
    let i = 0;
    i < data.length;
    i++
  ) {

    const currentComputerNo =
      String(data[i][3] || "")
        .trim()
        .toLowerCase();


    if (
      currentComputerNo ===
      searchComputerNo
    ) {

      return computerRowToObject(
        data[i]
      );

    }

  }


  return null;

}


/*************************************************
 * GET COMPUTER MANAGEMENT STATISTICS
 *************************************************/
function getComputerStats(
  sessionToken,
  filter
) {

  filter = filter || {};


  const computers =
    getComputerList(
      sessionToken,
      filter
    );


  const buildings =
    new Set();


  const labs =
    new Set();


  const users =
    new Set();


let printerCount = 0;


  computers.forEach(function(computer) {


    if (
      String(
        computer.buildingName
      ).trim()
    ) {

      buildings.add(
        computer.buildingName
      );

    }


    if (
      String(
        computer.labRoomName
      ).trim()
    ) {

      labs.add(
        computer.labRoomName
      );

    }


    if (
      String(
        computer.userName
      ).trim()
    ) {

      users.add(
        computer.userName
      );

    }


if (
  String(
    computer.printerBrand
  ).trim()
) {

  printerCount++;

}

  });


  return {

    totalComputers:
      computers.length,

    totalBuildings:
      buildings.size,

    totalLabsRooms:
      labs.size,

    assignedUsers:
      users.size,

printers:
  printerCount

  };

}


/*************************************************
 * GET COMPUTER FILTER OPTIONS
 *
 * Returns only values the logged-in user
 * is allowed to see.
 *************************************************/
function getComputerFilters(
  sessionToken
) {

  const data =
    getFilteredComputerData(
      sessionToken
    );


  const buildingSet =
    new Set();

  const labRoomSet =
    new Set();

  const floorSet =
    new Set();

  const assetTypeSet =
    new Set();

  const osSet =
    new Set();

  const cpuBrandSet =
    new Set();

  const processorSet =
    new Set();

  const ramSet =
    new Set();

  const primaryStorageSet =
    new Set();

  const graphicCardSet =
    new Set();

  const userSet =
    new Set();


  data.forEach(function(row) {


    if (row[1])
      buildingSet.add(
        String(row[1]).trim()
      );


    if (row[2])
      labRoomSet.add(
        String(row[2]).trim()
      );


    if (row[24])
      floorSet.add(
        String(row[24]).trim()
      );


    if (row[25])
      assetTypeSet.add(
        String(row[25]).trim()
      );


    if (row[14])
      osSet.add(
        String(row[14]).trim()
      );


    if (row[4])
      cpuBrandSet.add(
        String(row[4]).trim()
      );


    if (row[9])
      processorSet.add(
        String(row[9]).trim()
      );


    if (row[11])
      ramSet.add(
        String(row[11]).trim()
      );


    if (row[10])
      primaryStorageSet.add(
        String(row[10]).trim()
      );


    if (row[12])
      graphicCardSet.add(
        String(row[12]).trim()
      );


    if (row[20])
      userSet.add(
        String(row[20]).trim()
      );

  });


  return {

    buildings:
      Array.from(
        buildingSet
      ).sort(),

    labRooms:
      Array.from(
        labRoomSet
      ).sort(),

    floors:
      Array.from(
        floorSet
      ).sort(),

    assetTypes:
      Array.from(
        assetTypeSet
      ).sort(),

    operatingSystems:
      Array.from(
        osSet
      ).sort(),

    cpuBrands:
      Array.from(
        cpuBrandSet
      ).sort(),

    processors:
      Array.from(
        processorSet
      ).sort(),

    ramOptions:
      Array.from(
        ramSet
      ).sort(),

    primaryStorageOptions:
      Array.from(
        primaryStorageSet
      ).sort(),

    graphicCards:
      Array.from(
        graphicCardSet
      ).sort(),

    users:
      Array.from(
        userSet
      ).sort()

  };

}


/*************************************************
 * BUILDING-WISE COMPUTER CHART
 *************************************************/
function getComputerBuildingChart(
  sessionToken,
  filter
) {

  const computers =
    getComputerList(
      sessionToken,
      filter || {}
    );


  const result = {};


  computers.forEach(function(computer) {

    const building =
      String(
        computer.buildingName || ""
      ).trim();


    if (!building) {
      return;
    }


    if (!result[building]) {

      result[building] = {

        building:
          building,

        count:
          0

      };

    }


    result[building].count++;

  });


  return Object.values(
    result
  )
    .sort(function(a, b) {

      return b.count - a.count;

    });

}


/*************************************************
 * OS DISTRIBUTION CHART
 *************************************************/
function getOSDistributionChart(
  sessionToken,
  filter
) {

  const computers =
    getComputerList(
      sessionToken,
      filter || {}
    );


  const result = {};


  computers.forEach(function(computer) {

    let os =
      String(
        computer.os || ""
      ).trim();


    if (!os) {
      os = "Unknown";
    }


    if (!result[os]) {

      result[os] = {

        os: os,

        count: 0

      };

    }


    result[os].count++;

  });


  return Object.values(
    result
  )
    .sort(function(a, b) {

      return b.count - a.count;

    });

}


/*************************************************
 * RAM DISTRIBUTION CHART
 *************************************************/
function getRAMDistributionChart(
  sessionToken,
  filter
) {

  const computers =
    getComputerList(
      sessionToken,
      filter || {}
    );


  const result = {};


  computers.forEach(function(computer) {

    let ram =
      String(
        computer.ram || ""
      ).trim();


    if (!ram) {
      ram = "Unknown";
    }


    if (!result[ram]) {

      result[ram] = {

        ram: ram,

        count: 0

      };

    }


    result[ram].count++;

  });


  return Object.values(
    result
  );

}


/*************************************************
 * STORAGE DISTRIBUTION CHART
 *************************************************/
function getStorageDistributionChart(
  sessionToken,
  filter
) {

  const computers =
    getComputerList(
      sessionToken,
      filter || {}
    );


  const result = {};


  computers.forEach(function(computer) {

    let storage =
      String(
        computer.primaryStorage || ""
      ).trim();


    if (!storage) {
      storage = "Unknown";
    }


    if (!result[storage]) {

      result[storage] = {

        storage: storage,

        count: 0

      };

    }


    result[storage].count++;

  });


  return Object.values(
    result
  );

}


/*************************************************
 * GRAPHIC CARD DISTRIBUTION CHART
 *************************************************/
function getGraphicCardDistributionChart(
  sessionToken,
  filter
) {

  const computers =
    getComputerList(
      sessionToken,
      filter || {}
    );


  const result = {};


  computers.forEach(function(computer) {

    let graphicCard =
      String(
        computer.graphicCard || ""
      ).trim();


    if (!graphicCard) {
      graphicCard = "Unknown";
    }


    if (!result[graphicCard]) {

      result[graphicCard] = {

        graphicCard:
          graphicCard,

        count:
          0

      };

    }


    result[graphicCard].count++;

  });


  return Object.values(
    result
  )
    .sort(function(a, b) {

      return b.count - a.count;

    });

}


/*************************************************
 * GET CURRENT COMPUTER ACCESS INFORMATION
 *
 * Useful for frontend debugging
 *************************************************/
function getComputerAccessInfo(
  sessionToken
) {

  const session =
    getComputerUserSession(
      sessionToken
    );


  const access =
    getAssignedBuildings(
      sessionToken
    );


  return {

    userId:
      session.userId,

    name:
      session.name,

    email:
      session.email,

    role:
      session.role,

    isAdmin:
      access.isAdmin,

    assignedBuildings:
      access.buildings

  };

}


/*************************************************
 * TEST FUNCTION
 *
 * Run manually from Apps Script.
 *
 * Replace the token with an actual
 * session token from User_Sessions.
 *************************************************/
function testComputerManagement() {

  const token =
    "PASTE_VALID_SESSION_TOKEN_HERE";


  const access =
    getComputerAccessInfo(
      token
    );


  Logger.log(
    "ACCESS INFO:"
  );

  Logger.log(
    JSON.stringify(
      access,
      null,
      2
    )
  );


  const computers =
    getFilteredComputerData(
      token
    );


  Logger.log(
    "TOTAL COMPUTERS = " +
    computers.length
  );


  const stats =
    getComputerStats(
      token,
      {}
    );


  Logger.log(
    JSON.stringify(
      stats,
      null,
      2
    )
  );

}
function testComputerAccess() {

  const sheet = getSessionSheet();
  const data = sheet.getDataRange().getValues();

  // Get latest session token
  const lastRow = data[data.length - 1];

  const token = lastRow[0];

  Logger.log("Testing Token: " + token);

  const result = getComputerAccessInfo(token);

  Logger.log(JSON.stringify(result, null, 2));
}
/*************************************************
 * TEST COMPUTER DATA
 *************************************************/
function testComputerData() {

  const sheet = getComputerMasterSheet();

  const data = sheet.getDataRange().getValues();

  Logger.log("Sheet Name: " + sheet.getName());
  Logger.log("Total Rows Including Header: " + data.length);

  // Show headers
  Logger.log("HEADERS:");
  Logger.log(JSON.stringify(data[0]));

  // Show first computer record
  if (data.length > 1) {

    Logger.log("FIRST COMPUTER:");
    Logger.log(JSON.stringify(data[1]));

  }

}
function testComputerList() {

  const sheet = getSessionSheet();

  const data =
    sheet.getDataRange()
      .getValues();

  const lastRow =
    data[data.length - 1];

  const token =
    lastRow[0];

  const result =
    getComputerList(
      token,
      {}
    );

  Logger.log(
    "TOTAL COMPUTERS: " +
    result.length
  );

  Logger.log(
    JSON.stringify(
      result.slice(0, 3),
      null,
      2
    )
  );

}
function testComputerStats() {

  const sheet = getSessionSheet();
  const data = sheet.getDataRange().getValues();

  const token = data[data.length - 1][0];

  const result = getComputerStats(
    token,
    {}
  );

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

}
function testComputerFilters() {

  const sheet = getSessionSheet();
  const data = sheet.getDataRange().getValues();

  const token = data[data.length - 1][0];

  const result = getComputerFilters(token);

  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );

}
function testComputerCharts() {

  const sheet = getSessionSheet();
  const data = sheet.getDataRange().getValues();

  const token = data[data.length - 1][0];

  Logger.log("BUILDING CHART:");
  Logger.log(
    JSON.stringify(
      getComputerBuildingChart(token, {}),
      null,
      2
    )
  );

  Logger.log("OS CHART:");
  Logger.log(
    JSON.stringify(
      getOSDistributionChart(token, {}),
      null,
      2
    )
  );

}
