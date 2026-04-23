// Config.gs
const CONFIG = {
  TABLES: {
    TASKS: {
      NAME: "Tasks",
      COLUMNS: {
        PID: 0,
				ROOMID: 1,
        TASKID: 2,
				ROOMNAME: 3,
        TASKNAME: 4,
        VALUE: 5,
        UNIT: 6,
        PRICE: 7,
        COST: 8,
        NOTE: 9
      }
    },
    EQUIPMENT: {
      NAME: "Equipment",
      COLUMNS: {
        PID: 0,
        ITEMID: 1,
        ROOMID: 2,
        TASKID: 3,
        ROOMNAME: 4,
        TASKNAME: 5,
        ITEM: 6,
        VALUE: 7,
        UNIT: 8,
        PRICE: 9,
        COST: 10,
        NOTE: 11
      }
    },
		DETAILS: {
			NAME: "Details",
			COLUMNS: {
				PID: 0,
				ADDRESS1: 1,
				ADDRESS2: 2,
				CITY: 3,
				STATE: 4,
				ZIP: 5,
				COUNTRY: 6,
				FIRSTNAME: 7,
				LASTNAME: 8,
				EMAIL: 9,
				PHONE: 10
			}
    },
		DATES: {
			NAME: "Dates",
			COLUMNS: {
				PID: 0,
				LOSS: 1,
				DUE: 2,
				CONTACTED: 3,
				ASSIGNED: 4,
				INSPECTED: 5,
				ESTIMATED: 6,
				STARTED: 7,
				FINISHED: 8,
				INVOICED: 9,
				APPROVED: 10,
				PAID: 11
			}
		},
		ASSIGNMENTS: {
			NAME: "Assignments",
			COLUMNS: {
				PID: 0,
				ASSIGNMENTID: 1,
				ROLENAME: 2,
				FIRSTNAME: 3,
				LASTNAME: 4,
				MIDDLENAME: 5,
				EMAIL: 6,
				PHONE: 7,
				COMPANYCODE: 8
			}
		},
		CONTACTS: {
			NAME: "Contacts",
			COLUMNS: {
				COMPANYNAME: 0,
				COMPANYCODE: 1,
				ADDRESS1: 2,
				ADDRESS2: 3,
				CITY: 4,
				ZIP: 5,
				STATE: 6,
				COUNTRY: 7,
				FIRSTNAME: 8,
				LASTNAME: 9,
				MIDDLENAME: 10,
				ROLE: 11,
				EMAIL: 12,
				PHONE: 13,
				NOTES: 14
			}
		},
		MATERIALS: {
			NAME: "Materials",
			COLUMNS: {
				PID: 0,
				ROOMID: 1,
				TASKID: 2,
				ROOMNAME: 3,
				TASKNAME: 4,
				ITEMNAME: 5,
				VALUE: 6,
				UNIT: 7,
				PRICE: 8,
				COST: 9,
				NOTE: 10
			}
		},
		ROOMS: {
      NAME: "Rooms",
      COLUMNS: {
        PID: 0,
				ROOMNAME: 1,
				ROOMNUMBER: 2,
				LENGTH: 3,
				WIDTH: 4,
				HEIGHT: 5,
				ROOMID: 6,
				LENGTHUNIT: 7,
				WIDTHUNIT: 8,
				HEIGHTUNIT: 9
      }
    },
  },
  PAGES: {
    EQUIPMENT: {
      template: 'Page_Equipment',
      btn: 'Add Equipment',
      action: 'openEquipmentModal()',
      handler: 'renderEquipment',
      server: 'getEquipmentData',   
      icon: 'bi bi-plus-lg',
      showBtn: true,
      scripts: ['JS_Equipment', 'JS_Equipment_Modal'],
      title: 'Equipment'
    },
    TASKS: {
      template: 'Page_Tasks',
      btn: 'Add Task',
      action: 'openTaskModal()',
      handler: 'showTasks',
      server: 'getTasksData',   
      icon: 'bi bi-plus-lg',
      showBtn: true,
      scripts: ['JS_Tasks', 'JS_Tasks_Modal'],
      title: 'Tasks'
    },
	DETAILS: {
      template: 'Page_Tasks',
      btn: 'Edit Detail',
      action: 'openDetailModal()',
      handler: 'renderDetails',
      server: 'getDetailsData',   
      icon: 'bi bi-pencil',
      showBtn: true,
      scripts: ['JS_Tasks', 'JS_Tasks_Modal'],
      title: 'Details'
    }
  },
  FORMAT: {
    CURRENCY: (val) => {
      const num = Number(val) || 0;
      return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
      }).format(num);
    }
  }
};