const STORAGE_KEY = "noxxus.todo.tasks";

let tasks = loadTasks();

function loadTasks() {
  try {
    return JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
  } catch {
    return [];
  }
}

function saveTasks() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks));
}

function percent(done, total) {
  if (!total) return 0;
  return Math.round((done / total) * 100);
}

function totals() {
  return tasks.reduce(
    (base, task) => {
      base.total += task.quantity;
      base.done += task.done;
      return base;
    },
    { total: 0, done: 0 },
  );
}

function closeBrandToolbar() {
  const button = document.querySelector("#brandMenuButton");
  const toolbar = document.querySelector("#brandToolbar");
  const dim = document.querySelector("#pageDim");
  toolbar.hidden = true;
  dim.hidden = true;
  button.setAttribute("aria-expanded", "false");
}

function toggleBrandToolbar() {
  const button = document.querySelector("#brandMenuButton");
  const toolbar = document.querySelector("#brandToolbar");
  const dim = document.querySelector("#pageDim");
  const shouldOpen = toolbar.hidden;
  toolbar.hidden = !shouldOpen;
  dim.hidden = !shouldOpen;
  button.setAttribute("aria-expanded", String(shouldOpen));
}

function changeDone(id, amount) {
  tasks = tasks.map((task) => {
    if (task.id !== id) return task;
    const done = Math.max(0, Math.min(task.quantity, task.done + amount));
    return { ...task, done };
  });
  saveTasks();
  render();
}

function removeTask(id) {
  tasks = tasks.filter((task) => task.id !== id);
  saveTasks();
  render();
}

function priorityClass(priority) {
  return priority.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function renderChart() {
  const total = totals();
  const value = percent(total.done, total.total);
  document.querySelector("#todoTotalLabel").textContent = `${total.done}/${total.total}`;
  document.querySelector("#todoPercent").textContent = `${value}%`;
  document.querySelector("#todoChart").style.setProperty("--progress", `${value}%`);
}

function renderList() {
  const list = document.querySelector("#todoList");
  document.querySelector("#todoListHint").textContent = `${tasks.length} tarefas`;
  if (!tasks.length) {
    list.innerHTML = `<p class="empty-state">Nenhuma tarefa cadastrada.</p>`;
    return;
  }

  list.innerHTML = tasks
    .map((task) => {
      const taskPercent = percent(task.done, task.quantity);
      return `<article class="todo-card">
        <div class="todo-card-head">
          <h3>${task.name}</h3>
          <span class="priority-pill ${priorityClass(task.priority)}">${task.priority}</span>
        </div>
        <div class="todo-meta">
          <span>${task.done}/${task.quantity} tarefas</span>
          <strong>${taskPercent}%</strong>
        </div>
        <div class="progress-bar" aria-label="Conclusao individual">
          <span style="width: ${taskPercent}%"></span>
        </div>
        <div class="todo-actions">
          <button class="ghost" data-minus="${task.id}" type="button">-1</button>
          <button data-plus="${task.id}" type="button">+1</button>
          <button class="danger" data-remove="${task.id}" type="button">Remover</button>
        </div>
      </article>`;
    })
    .join("");

  document.querySelectorAll("[data-minus]").forEach((button) => {
    button.addEventListener("click", () => changeDone(button.dataset.minus, -1));
  });
  document.querySelectorAll("[data-plus]").forEach((button) => {
    button.addEventListener("click", () => changeDone(button.dataset.plus, 1));
  });
  document.querySelectorAll("[data-remove]").forEach((button) => {
    button.addEventListener("click", () => removeTask(button.dataset.remove));
  });
}

function render() {
  renderChart();
  renderList();
}

document.querySelector("#todoForm").addEventListener("submit", (event) => {
  event.preventDefault();
  const nameInput = document.querySelector("#taskName");
  const quantityInput = document.querySelector("#taskQuantity");
  const quantity = Math.max(1, Number(quantityInput.value || 1));
  tasks.unshift({
    id: `${Date.now()}`,
    name: nameInput.value.trim(),
    priority: document.querySelector("#taskPriority").value,
    quantity,
    done: 0,
  });
  saveTasks();
  event.target.reset();
  quantityInput.value = "1";
  nameInput.focus();
  render();
});

document.querySelector("#brandMenuButton").addEventListener("click", (event) => {
  event.stopPropagation();
  toggleBrandToolbar();
});
document.addEventListener("click", (event) => {
  if (!event.target.closest(".brand-menu")) closeBrandToolbar();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeBrandToolbar();
});

render();
