/* =========================================================
   WORK HANDOVER BACKEND
========================================================= */


/* =========================================================
   GET SPREADSHEET
========================================================= */

function getWorkHandoverSpreadsheet() {

  return SpreadsheetApp.openById(SPREADSHEET_ID);

}


/* =========================================================
   GET SHEETS
========================================================= */

function getWorkHandoverSheets() {

  const ss = getWorkHandoverSpreadsheet();

  return {

    workload:
      ss.getSheetByName("WorkLoad"),

    buildingMap:
      ss.getSheetByName("Building_Technician_Map"),

    users:
      ss.getSheetByName("Users")

  };

}


/* =========================================================
   GET CURRENT USER
========================================================= */

function getCurrentWorkHandoverUser(token) {

  if (!token) {

    throw new Error(
      "Session token is missing."
    );

  }


  const user =
    getUserSession(token);


  if (!user) {

    throw new Error(
      "Session expired. Please login again."
    );

  }


  return user;

}


/* =========================================================
   GET FORM DATA
========================================================= */

function getWorkHandoverFormData(token) {

  const user =
    getCurrentWorkHandoverUser(token);


  const sheets =
    getWorkHandoverSheets();


  if (!sheets.buildingMap) {

    throw new Error(
      "Building_Technician_Map sheet not found."
    );

  }


  if (!sheets.users) {

    throw new Error(
      "Users sheet not found."
    );

  }


  const buildingData =
    sheets.buildingMap
      .getDataRange()
      .getValues();


  const userData =
    sheets.users
      .getDataRange()
      .getValues();


  /*
     -------------------------------------------------------
     BUILDING MAP HEADER
     -------------------------------------------------------
  */

  const buildingHeaders =
    buildingData[0].map(function(header) {

      return String(header)
        .trim();

    });


  const buildingNameIndex =
    buildingHeaders.indexOf(
      "Building Name"
    );


  const technicianNameIndex =
    buildingHeaders.indexOf(
      "Technician Name"
    );


  const activeStatusIndex =
    buildingHeaders.indexOf(
      "Active Status"
    );


  if (
    buildingNameIndex === -1 ||
    technicianNameIndex === -1
  ) {

    throw new Error(
      "Building_Technician_Map headers are incorrect."
    );

  }


  /*
     -------------------------------------------------------
     GET BUILDINGS ASSIGNED TO CURRENT TECHNICIAN
     -------------------------------------------------------
  */

  const buildings = [];


  for (
    let i = 1;
    i < buildingData.length;
    i++
  ) {

    const row =
      buildingData[i];


    const building =
      String(
        row[buildingNameIndex] || ""
      ).trim();


    const technician =
      String(
        row[technicianNameIndex] || ""
      ).trim();


    const active =
      activeStatusIndex !== -1
      ?
      String(
        row[activeStatusIndex] || ""
      ).trim().toUpperCase()
      :
      "YES";


    if (
      building &&
      technician.toLowerCase() ===
        String(user.name)
          .trim()
          .toLowerCase() &&
      active === "YES"
    ) {

      /*
         Prevent duplicate buildings.
      */

      const alreadyExists =
        buildings.some(function(item) {

          return item.building
            .toLowerCase()
            ===
            building.toLowerCase();

        });


      if (!alreadyExists) {

        buildings.push({

          building:
            building

        });

      }

    }

  }


  /*
     -------------------------------------------------------
     USERS HEADER
     -------------------------------------------------------
  */

  const userHeaders =
    userData[0].map(function(header) {

      return String(header)
        .trim();

    });


  const nameIndex =
    userHeaders.indexOf(
      "Name"
    );


  const emailIndex =
    userHeaders.indexOf(
      "Email"
    );


  const mobileIndex =
    userHeaders.indexOf(
      "Mobile"
    );


  const roleIndex =
    userHeaders.indexOf(
      "Role"
    );


  const statusIndex =
    userHeaders.indexOf(
      "Status"
    );


  if (
    nameIndex === -1 ||
    emailIndex === -1 ||
    mobileIndex === -1 ||
    roleIndex === -1 ||
    statusIndex === -1
  ) {

    throw new Error(
      "Users sheet headers are incorrect."
    );

  }


  /*
     -------------------------------------------------------
     GET ACTIVE TECHNICIANS
     -------------------------------------------------------
  */

  const technicians = [];


  for (
    let i = 1;
    i < userData.length;
    i++
  ) {

    const row =
      userData[i];


    const name =
      String(
        row[nameIndex] || ""
      ).trim();


    const email =
      String(
        row[emailIndex] || ""
      ).trim();


    const mobile =
      String(
        row[mobileIndex] || ""
      ).trim();


    const role =
      String(
        row[roleIndex] || ""
      ).trim()
      .toUpperCase();


    const status =
      String(
        row[statusIndex] || ""
      ).trim()
      .toUpperCase();


    if (
      name &&
      role === "TECHNICIAN" &&
      status === "YES"
    ) {

      /*
         Don't show the currently logged-in technician
         as the override technician.
      */

      if (
        name.toLowerCase() ===
        String(user.name)
          .trim()
          .toLowerCase()
      ) {

        continue;

      }


      technicians.push({

        name:
          name,

        email:
          email,

        mobile:
          mobile

      });

    }

  }


  return {

    user: {

      name:
        user.name,

      email:
        user.email,

      mobile:
        user.mobile,

      role:
        user.role

    },

    buildings:
      buildings,

    technicians:
      technicians

  };

}


/* =========================================================
   GET TECHNICIAN DETAILS
========================================================= */

function getWorkHandoverTechnicianDetails(
  token,
  technicianName
) {

  const user =
    getCurrentWorkHandoverUser(token);


  /*
     Only authenticated users can request
     technician details.
  */


  if (!technicianName) {

    throw new Error(
      "Technician name is required."
    );

  }


  const sheets =
    getWorkHandoverSheets();


  const data =
    sheets.users
      .getDataRange()
      .getValues();


  const headers =
    data[0].map(function(header) {

      return String(header)
        .trim();

    });


  const nameIndex =
    headers.indexOf("Name");


  const emailIndex =
    headers.indexOf("Email");


  const mobileIndex =
    headers.indexOf("Mobile");


  const roleIndex =
    headers.indexOf("Role");


  const statusIndex =
    headers.indexOf("Status");


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const name =
      String(
        row[nameIndex] || ""
      ).trim();


    if (
      name.toLowerCase() ===
      String(technicianName)
        .trim()
        .toLowerCase()
    ) {

      const role =
        String(
          row[roleIndex] || ""
        ).trim()
        .toUpperCase();


      const status =
        String(
          row[statusIndex] || ""
        ).trim()
        .toUpperCase();


      if (
        role !== "TECHNICIAN" ||
        status !== "YES"
      ) {

        throw new Error(
          "Selected technician is not active."
        );

      }


      /*
         Don't allow current technician
         to select themselves.
      */

      if (
        name.toLowerCase() ===
        String(user.name)
          .trim()
          .toLowerCase()
      ) {

        throw new Error(
          "You cannot select yourself."
        );

      }


      return {

        name:
          name,

        email:
          String(
            row[emailIndex] || ""
          ).trim(),

        mobile:
          String(
            row[mobileIndex] || ""
          ).trim()

      };

    }

  }


  throw new Error(
    "Technician not found."
  );

}


/* =========================================================
   CREATE WORK HANDOVER
========================================================= */

function createWorkHandover(
  token,
  data
) {

  const user =
    getCurrentWorkHandoverUser(token);


  if (!data) {

    throw new Error(
      "Handover data is missing."
    );

  }


  const sheets =
    getWorkHandoverSheets();


  if (!sheets.workload) {

    throw new Error(
      "WorkLoad sheet not found."
    );

  }


  /*
     -------------------------------------------------------
     BASIC VALIDATION
     -------------------------------------------------------
  */

  const building =
    String(
      data.building || ""
    ).trim();


  const overrideTechnician =
    String(
      data.overrideTechnician || ""
    ).trim();


  const startDate =
    String(
      data.startDate || ""
    ).trim();


  const endDate =
    String(
      data.endDate || ""
    ).trim();


  const reason =
    String(
      data.reason || ""
    ).trim();


  if (!building) {

    throw new Error(
      "Building is required."
    );

  }


  if (!overrideTechnician) {

    throw new Error(
      "Override technician is required."
    );

  }


  if (!startDate) {

    throw new Error(
      "Start date is required."
    );

  }


  if (!endDate) {

    throw new Error(
      "End date is required."
    );

  }


  /*
     -------------------------------------------------------
     ORIGINAL TECHNICIAN MUST BE LOGGED-IN USER
     -------------------------------------------------------
  */

  const originalTechnician =
    String(user.name || "")
      .trim();


  if (!originalTechnician) {

    throw new Error(
      "Unable to identify current technician."
    );

  }


  /*
     -------------------------------------------------------
     PREVENT SELF ASSIGNMENT
     -------------------------------------------------------
  */

  if (
    originalTechnician
      .toLowerCase()
    ===
    overrideTechnician
      .toLowerCase()
  ) {

    throw new Error(
      "You cannot assign the handover to yourself."
    );

  }


  /*
     -------------------------------------------------------
     VERIFY BUILDING BELONGS TO CURRENT TECHNICIAN
     -------------------------------------------------------
  */

  verifyTechnicianBuilding(
    originalTechnician,
    building
  );


  /*
     -------------------------------------------------------
     VERIFY OVERRIDE TECHNICIAN
     -------------------------------------------------------
  */

  const technician =
    getWorkHandoverTechnicianDetails(
      token,
      overrideTechnician
    );


  /*
     -------------------------------------------------------
     DATE VALIDATION
     -------------------------------------------------------
  */

  const start =
    new Date(startDate);


  const end =
    new Date(endDate);


  if (
    isNaN(start.getTime()) ||
    isNaN(end.getTime())
  ) {

    throw new Error(
      "Invalid handover date."
    );

  }


  if (end < start) {

    throw new Error(
      "End Date cannot be earlier than Start Date."
    );

  }


  /*
     -------------------------------------------------------
     CHECK EXISTING ACTIVE HANDOVER
     -------------------------------------------------------
  */

  const conflict =
    checkWorkHandoverConflict(
      building,
      start,
      end
    );


  if (conflict) {

    throw new Error(
      "An active handover already exists for this building during the selected dates."
    );

  }


  /*
     -------------------------------------------------------
     WORKLOAD HEADERS
     -------------------------------------------------------
  */

  const workloadData =
    sheets.workload
      .getDataRange()
      .getValues();


  const headers =
    workloadData[0].map(function(header) {

      return String(header)
        .trim();

    });


  const requiredHeaders = [

    "Timestamp",

    "Building Name",

    "Original Technician",

    "Override Technician",

    "Override Technician Mobile",

    "Override Start Date",

    "Override End Date",

    "Override Reason",

    "Active",

    "Override Technician Mail Id"

  ];


  requiredHeaders.forEach(function(header) {

    if (
      headers.indexOf(header) === -1
    ) {

      throw new Error(
        "WorkLoad sheet is missing header: " +
        header
      );

    }

  });


  /*
     -------------------------------------------------------
     PREPARE NEW ROW
     -------------------------------------------------------
  */

  const newRow =
    new Array(headers.length)
      .fill("");


  setWorkloadValue(
    newRow,
    headers,
    "Timestamp",
    new Date()
  );


  setWorkloadValue(
    newRow,
    headers,
    "Building Name",
    building
  );


  setWorkloadValue(
    newRow,
    headers,
    "Original Technician",
    originalTechnician
  );


  setWorkloadValue(
    newRow,
    headers,
    "Override Technician",
    technician.name
  );


  setWorkloadValue(
    newRow,
    headers,
    "Override Technician Mobile",
    technician.mobile
  );


  setWorkloadValue(
    newRow,
    headers,
    "Override Start Date",
    start
  );


  setWorkloadValue(
    newRow,
    headers,
    "Override End Date",
    end
  );


  setWorkloadValue(
    newRow,
    headers,
    "Override Reason",
    reason
  );


/*
   Automatically calculate status
*/

const handoverStatus =
  getWorkHandoverStatus(
    start,
    end
  );


setWorkloadValue(
  newRow,
  headers,
  "Active",
  handoverStatus
);


  setWorkloadValue(
    newRow,
    headers,
    "Override Technician Mail Id",
    technician.email
  );


  /*
     -------------------------------------------------------
     SAVE
     -------------------------------------------------------
  */

  sheets.workload.appendRow(
    newRow
  );


  return {

    success:
      true,

    message:
      "Work handover created successfully.",

    record: {

      building:
        building,

      originalTechnician:
        originalTechnician,

      overrideTechnician:
        technician.name,

      startDate:
        startDate,

      endDate:
        endDate,

active:
  handoverStatus

    }

  };

}


/* =========================================================
   SET WORKLOAD VALUE
========================================================= */

function setWorkloadValue(
  row,
  headers,
  headerName,
  value
) {

  const index =
    headers.indexOf(headerName);


  if (index !== -1) {

    row[index] =
      value;

  }

}


/* =========================================================
   VERIFY TECHNICIAN BUILDING
========================================================= */

function verifyTechnicianBuilding(
  technicianName,
  buildingName
) {

  const sheets =
    getWorkHandoverSheets();


  const data =
    sheets.buildingMap
      .getDataRange()
      .getValues();


  const headers =
    data[0].map(function(header) {

      return String(header)
        .trim();

    });


  const buildingIndex =
    headers.indexOf(
      "Building Name"
    );


  const technicianIndex =
    headers.indexOf(
      "Technician Name"
    );


  const activeIndex =
    headers.indexOf(
      "Active Status"
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const building =
      String(
        row[buildingIndex] || ""
      ).trim();


    const technician =
      String(
        row[technicianIndex] || ""
      ).trim();


    const active =
      activeIndex !== -1
      ?
      String(
        row[activeIndex] || ""
      ).trim().toUpperCase()
      :
      "YES";


    if (
      building.toLowerCase() ===
        buildingName.toLowerCase() &&

      technician.toLowerCase() ===
        technicianName.toLowerCase() &&

      active === "YES"
    ) {

      return true;

    }

  }


  throw new Error(
    "This building is not assigned to you."
  );

}


/* =========================================================
   CHECK HANDOVER CONFLICT
========================================================= */

/* =========================================================
   CHECK HANDOVER CONFLICT
========================================================= */

function checkWorkHandoverConflict(
  building,
  startDate,
  endDate
) {

  const sheets =
    getWorkHandoverSheets();


  if (!sheets.workload) {

    return false;

  }


  const data =
    sheets.workload
      .getDataRange()
      .getValues();


  if (data.length <= 1) {

    return false;

  }


  const headers =
    data[0].map(function(header) {

      return String(header)
        .trim();

    });


  const buildingIndex =
    headers.indexOf(
      "Building Name"
    );


  const startIndex =
    headers.indexOf(
      "Override Start Date"
    );


  const endIndex =
    headers.indexOf(
      "Override End Date"
    );


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const rowBuilding =
      String(
        row[buildingIndex] || ""
      ).trim();


    /*
       Check same building
    */

    if (
      rowBuilding.toLowerCase() !==
      building.toLowerCase()
    ) {

      continue;

    }


    /*
       Read existing handover dates
    */

    const existingStart =
      new Date(
        row[startIndex]
      );


    const existingEnd =
      new Date(
        row[endIndex]
      );


    /*
       Ignore invalid dates
    */

    if (
      isNaN(existingStart.getTime()) ||
      isNaN(existingEnd.getTime())
    ) {

      continue;

    }


    /*
       Calculate actual status
       from the dates.
    */

    const existingStatus =
      getWorkHandoverStatus(
        existingStart,
        existingEnd
      );


    /*
       Only ACTIVE handovers
       can create a conflict.
    */

    if (
      existingStatus !== "ACTIVE"
    ) {

      continue;

    }


    /*
       Date overlap condition:

       New Start <= Existing End
       AND
       New End >= Existing Start
    */

    if (
      startDate <= existingEnd &&
      endDate >= existingStart
    ) {

      return true;

    }

  }


  return false;

}


/* =========================================================
   GET WORK HANDOVER RECORDS
========================================================= */

/* =========================================================
   GET WORK HANDOVER RECORDS
========================================================= */

function getWorkHandoverRecords(token) {

  const user =
    getCurrentWorkHandoverUser(token);


  const sheets =
    getWorkHandoverSheets();


  if (!sheets.workload) {

    throw new Error(
      "WorkLoad sheet not found."
    );

  }


  const sheet =
    sheets.workload;


  const data =
    sheet
      .getDataRange()
      .getValues();


  if (data.length <= 1) {

    return [];

  }


  const headers =
    data[0].map(function(header) {

      return String(header)
        .trim();

    });


  const timestampIndex =
    headers.indexOf(
      "Timestamp"
    );


  const buildingIndex =
    headers.indexOf(
      "Building Name"
    );


  const originalIndex =
    headers.indexOf(
      "Original Technician"
    );


  const overrideIndex =
    headers.indexOf(
      "Override Technician"
    );


  const startIndex =
    headers.indexOf(
      "Override Start Date"
    );


  const endIndex =
    headers.indexOf(
      "Override End Date"
    );


  const activeIndex =
    headers.indexOf(
      "Active"
    );


  const reasonIndex =
    headers.indexOf(
      "Override Reason"
    );


  const records = [];


  const role =
    String(user.role || "")
      .trim()
      .toUpperCase();


  const currentUserName =
    String(user.name || "")
      .trim()
      .toLowerCase();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const original =
      String(
        row[originalIndex] || ""
      ).trim();


    const override =
      String(
        row[overrideIndex] || ""
      ).trim();


    /*
       ---------------------------------------------------
       USER VISIBILITY
       ---------------------------------------------------

       ADMIN:
       See all handovers.

       TECHNICIAN:
       See handovers:

       1. Created by him
       OR
       2. Assigned to him
    */

    let visible = false;


    if (role === "ADMIN") {

      visible = true;

    }
    else {

      if (
        original.toLowerCase() ===
        currentUserName
      ) {

        visible = true;

      }


      if (
        override.toLowerCase() ===
        currentUserName
      ) {

        visible = true;

      }

    }


    if (!visible) {

      continue;

    }


    /*
       ---------------------------------------------------
       GET DATES
       ---------------------------------------------------
    */

    const startDate =
      new Date(
        row[startIndex]
      );


    const endDate =
      new Date(
        row[endIndex]
      );


    /*
       ---------------------------------------------------
       CALCULATE REAL STATUS
       ---------------------------------------------------
    */

    const status =
      getWorkHandoverStatus(
        startDate,
        endDate
      );


    /*
       ---------------------------------------------------
       UPDATE ACTIVE COLUMN
       ---------------------------------------------------

       This keeps the Google Sheet synchronized.

       ACTIVE   -> ACTIVE
       UPCOMING -> UPCOMING
       EXPIRED  -> EXPIRED
    */

    if (activeIndex !== -1) {

      const sheetStatus =
        String(
          row[activeIndex] || ""
        )
        .trim()
        .toUpperCase();


      if (
        sheetStatus !== status
      ) {

        sheet
          .getRange(
            i + 1,
            activeIndex + 1
          )
          .setValue(status);

      }

    }


    /*
       ---------------------------------------------------
       ADD RECORD
       ---------------------------------------------------
    */

    records.push({

      timestamp:
        formatWorkHandoverDate(
          row[timestampIndex]
        ),

      building:
        row[buildingIndex] || "",

      originalTechnician:
        original,

      overrideTechnician:
        override,

      startDate:
        formatWorkHandoverDate(
          row[startIndex]
        ),

      endDate:
        formatWorkHandoverDate(
          row[endIndex]
        ),

      reason:
        row[reasonIndex] || "",

      active:
        status

    });

  }


  /*
     Newest first.
  */

  records.reverse();


  return records;

}


/* =========================================================
   FORMAT DATE
========================================================= */

function formatWorkHandoverDate(value) {

  if (!value) {

    return "";

  }


  if (
    Object.prototype.toString
      .call(value)
      === "[object Date]"
  ) {

    return Utilities.formatDate(

      value,

      Session.getScriptTimeZone(),

      "dd-MM-yyyy"

    );

  }


  const date =
    new Date(value);


  if (!isNaN(date.getTime())) {

    return Utilities.formatDate(

      date,

      Session.getScriptTimeZone(),

      "dd-MM-yyyy"

    );

  }


  return String(value);

}
/* =========================================================
   GET HANDOVER STATUS
========================================================= */

function getWorkHandoverStatus(startDate, endDate) {

  if (!startDate || !endDate) {

    return "EXPIRED";

  }

  const start =
    new Date(startDate);

  const end =
    new Date(endDate);

  if (
    isNaN(start.getTime()) ||
    isNaN(end.getTime())
  ) {

    return "EXPIRED";

  }


  /*
     Remove time from all dates.
  */

  start.setHours(0, 0, 0, 0);

  end.setHours(0, 0, 0, 0);


  const today =
    new Date();

  today.setHours(0, 0, 0, 0);


  /*
     Before Start Date
  */

  if (today < start) {

    return "UPCOMING";

  }


  /*
     Between Start Date and End Date
  */

  if (
    today >= start &&
    today <= end
  ) {

    return "ACTIVE";

  }


  /*
     After End Date
  */

  return "EXPIRED";

}
