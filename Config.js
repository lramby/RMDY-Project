// Config.gs
const CONFIG = {
  TABLES: {
    TASKS: {
      NAME: "Tasks",
      COLUMNS: {
        PID: 0,
        TASKID: 1,
        ROOMID: 2,
        TASK: 3,
        ROOMNAME: 4,
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
    }
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