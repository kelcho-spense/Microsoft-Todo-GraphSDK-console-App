// Create an authentication provider
const authProvider = {
    getAccessToken: async () => {
        // Call getToken in auth.js
        return await getToken();
    }
};
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });
//Get user info from Graph
async function getUser() {
    ensureScope('user.read');
    const DisplayName = await graphClient.api('/me').select('id,displayName').get();
    sessionStorage.setItem('displayName', DisplayName.displayName);
    return DisplayName;
}

async function run() {

    ensureScope('Tasks.ReadWrite');
    //1. ❇Task List Operations❇
    // get all task lists✔
    async function getTaskLists() {
        return await graphClient.api('/me/todo/lists').get();
    }
    //create a tasklist✔
    async function addTaskList(taskName) {
        const taskList = {
            displayName: taskName
        };
        return await graphClient.api('/me/todo/lists').post(taskList);
    }
    //delete a tasklist
    async function deleteTaskList(taskListId) {
        return await graphClient.api('/me/todo/lists/' + taskListId).delete();
    }
    //update a tasklist
    async function updateTaskList(taskListId, taskName) {
        const taskList = {
            displayName: taskName
        };
        return await graphClient.api('/me/todo/lists/' + taskListId).patch(taskList);
    }

    //2. ❇Task Operations❇

    //create a task in a task list
    async function createATask(taskListId, taskName) {
        const todoTask = {
            title: taskName
        };
        return await graphClient.api('/me/todo/lists/' + taskListId + '/tasks').post(todoTask);
    }

    //get all tasks in a task list
    async function getAllTasksInATaskList(taskListId) {
        return await graphClient.api('/me/todo/lists/' + taskListId + '/tasks').get();
    }

    // delete a single task in a task list
    async function deleteTaskInTaskList(taskListId, taskId) {
        return await graphClient.api('/me/todo/lists/' + taskListId + '/tasks/' + taskId).delete();
    }

    // update a single task in a task list
    async function updateTaskInTaskList(taskListId, taskId, taskName) {
        const todoTask = {
            title: taskName
        };
        return await graphClient.api('/me/todo/lists/' + taskListId + '/tasks/' + taskId).patch(todoTask);
    }


    //3.❇function calls for taskList Operations❇
    //1. get all task lists✔
    // const taskLists = await getTaskLists();
    // console.dir(taskLists);
    //2. add a task list✔
    // const addTask = await addTaskList("New Task List");
    // console.dir(addTask);
    //3. delete a task list✔
    // taskId = "AAMkAGFlNTNhY2U0LWQ0ZmUtNDMzOS05ZTBjLTczMDU4OWQ5NDM5NAAuAAAAAABHlovkJnSFQbO3LTz446G6AQBYcw3l7AL0RaThMm2-upfrAAARFIyHAAA=";
    // await deleteTaskList(taskId);
    // 4. update a task list✔
    // const updateTaskId = "AAMkAGFlNTNhY2U0LWQ0ZmUtNDMzOS05ZTBjLTczMDU4OWQ5NDM5NAAuAAAAAABHlovkJnSFQbO3LTz446G6AQBYcw3l7AL0RaThMm2-upfrAAARFIyEAAA=";
    // const updatedTaskList = await updateTaskList(updateTaskId, "Updated Task List");






    //4.❇function calls for task Operations❇
    //get all task lists✔

    // const taskLists = await getTaskLists();
    // console.dir(taskLists);

    // add task list✔
    // const addTask = await addTaskList("Task 1");
    // console.dir('The new Task: ' + addTask);
    // console.dir(taskLists);

    // add a task in a task list✔
    // let taskListId = "AQMkAGFlNTNhY2U0LWQ0ZmUtNDMzADktOWUwYy03MzA1ODlkOTQzOTQALgAAA0eWi_QmdIVBs7ctPPjjoboBAFhzDeXsAvRFpOEybb_6l_sAAAIBEgAAAA==";
    // let taskTitle = 'task 3'
    // const createdTaskInTaskList = await createATask(taskListId, taskTitle);
    // console.dir(createdTaskInTaskList);

    // get all tasks in a task list✔

    // const getTasksInTaskList = await getAllTasksInATaskList(taskListId);
    // console.dir(getTasksInTaskList);

    // delete a single task in a task list✔
    // const taskId = "AAMkAGFlNTNhY2U0LWQ0ZmUtNDMzOS05ZTBjLTczMDU4OWQ5NDM5NABGAAAAAABHlovkJnSFQbO3LTz446G6BwBYcw3l7AL0RaThMm2-upfrAAAAAAESAABYcw3l7AL0RaThMm2-upfrAAARFKYAAAA=";
    // await deleteTaskInTaskList(taskListId, taskId);
    // console.dir(getTasksInTaskList);

    // update a single task in a task list✔
    // const taskId = "AAMkAGFlNTNhY2U0LWQ0ZmUtNDMzOS05ZTBjLTczMDU4OWQ5NDM5NABGAAAAAABHlovkJnSFQbO3LTz446G6BwBYcw3l7AL0RaThMm2-upfrAAAAAAESAABYcw3l7AL0RaThMm2-upfrAAARFKX_AAA=";
    // const updatedTask = await updateTaskInTaskList(taskListId, taskId, 'updated task 1');

    // const updatedGetTasksInTaskList = await getAllTasksInATaskList(taskListId);
    // console.dir(updatedGetTasksInTaskList);






    // delete all task lists
    // const deleteAllTaskLists = await DeleteALLList();
    // console.dir(taskLists);

}







