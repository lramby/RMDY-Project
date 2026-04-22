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
