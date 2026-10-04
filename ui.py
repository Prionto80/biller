from nicegui import ui

todos = []

editing_task = None
edit_input = None

def add_todo():
    if todo_input.value:
        todos.append(
            {
                "task": todo_input.value,
                "done": False,
            }
        )
        todo_input.value = ''
        refresh_todo()

def delete_todo(todo):
    todos.remove(todo)
    refresh_todo()

def update_todo(todo, edit_input):
    global editing_task
    todo['task'] = edit_input.value
    editing_task = None
    refresh_todo()

def start_edit(todo):
    global editing_task
    editing_task = todo
    refresh_todo()


def refresh_todo():
    todo_container.clear()

    with todo_container:
        for todo in todos:
            with ui.row():

                if todo == editing_task:
                    
                    edit_input = ui.input(
                        value = todo['task']
                    )

                    ui.button(
                        'Update',
                        on_click=lambda t=todo: update_todo(t, edit_input)
                    )

                else:
                    checkbox = ui.checkbox(
                        todo['task'],
                        value=todo['done'],
                    )

                    checkbox.bind_value(todo, 'done')
                    checkbox.on(
                        'dblclick',
                        lambda e,t=todo: start_edit(t)
                    )

                    ui.button(
                        'Delete',
                        on_click=lambda t = todo: delete_todo(t)
                    )

todo_input = ui.input("New Task")
ui.button('Add Task', on_click=add_todo)

todo_container = ui.column()

ui.button('print done', on_click=lambda : print(todos))

ui.run()