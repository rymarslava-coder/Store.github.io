const ORDERS_KEY = 'order-catalog-orders'
const SCALE_KEY = 'order-catalog-font-scale'

const nomenclature = [
  { id: 'n-001', name: 'Бумага офисная А4', unit: 'пачка', priceTiers: [
    { minQty: 1, maxQty: 5, price: 450 }, { minQty: 6, maxQty: 20, price: 420 }, { minQty: 21, maxQty: null, price: 390 },
  ]},
  { id: 'n-002', name: 'Ручка шариковая синяя', unit: 'шт', priceTiers: [
    { minQty: 1, maxQty: 10, price: 35 }, { minQty: 11, maxQty: 50, price: 30 }, { minQty: 51, maxQty: null, price: 25 },
  ]},
  { id: 'n-003', name: 'Папка-регистратор', unit: 'шт', priceTiers: [
    { minQty: 1, maxQty: 3, price: 280 }, { minQty: 4, maxQty: 15, price: 250 }, { minQty: 16, maxQty: null, price: 220 },
  ]},
  { id: 'n-004', name: 'Степлер настольный', unit: 'шт', priceTiers: [
    { minQty: 1, maxQty: 2, price: 650 }, { minQty: 3, maxQty: 10, price: 590 }, { minQty: 11, maxQty: null, price: 540 },
  ]},
  { id: 'n-005', name: 'Картридж для принтера', unit: 'шт', priceTiers: [
    { minQty: 1, maxQty: 2, price: 3200 }, { minQty: 3, maxQty: 10, price: 2950 }, { minQty: 11, maxQty: null, price: 2700 },
  ]},
  { id: 'n-006', name: 'Блокнот А5', unit: 'шт', priceTiers: [
    { minQty: 1, maxQty: 10, price: 120 }, { minQty: 11, maxQty: 30, price: 105 }, { minQty: 31, maxQty: null, price: 90 },
  ]},
  { id: 'n-007', name: 'Маркер перманентный', unit: 'шт', priceTiers: [
    { minQty: 1, maxQty: 5, price: 85 }, { minQty: 6, maxQty: 20, price: 75 }, { minQty: 21, maxQty: null, price: 65 },
  ]},
  { id: 'n-008', name: 'Скотч упаковочный', unit: 'рулон', priceTiers: [
    { minQty: 1, maxQty: 5, price: 180 }, { minQty: 6, maxQty: 20, price: 160 }, { minQty: 21, maxQty: null, price: 140 },
  ]},
]

const categories = [
  { id: 'cat-office', name: 'Канцелярские товары', itemIds: ['n-001', 'n-002', 'n-003', 'n-004', 'n-006'] },
  { id: 'cat-supplies', name: 'Расходные материалы', itemIds: ['n-005', 'n-007', 'n-008'] },
]

const state = {
  view: 'catalog',
  cart: [],
  scale: parseFloat(localStorage.getItem(SCALE_KEY)) || 1.2,
  expanded: {},
  selected: {},
  quantity: {},
}

function getItem(id) { return nomenclature.find((i) => i.id === id) }

function getPrice(item, qty) {
  if (!item || qty <= 0) return 0
  const tier = item.priceTiers.find((t) => qty >= t.minQty && (t.maxQty === null || qty <= t.maxQty))
  return tier ? tier.price : item.priceTiers[item.priceTiers.length - 1].price
}

function formatPrice(v) {
  return new Intl.NumberFormat('ru-RU', { style: 'currency', currency: 'RUB', maximumFractionDigits: 0 }).format(v)
}

function showToast(msg) {
  const el = document.getElementById('toast')
  el.textContent = msg
  el.classList.remove('hidden')
  setTimeout(() => el.classList.add('hidden'), 3000)
}

function applyScale() {
  document.documentElement.style.setProperty('--font-scale', String(state.scale))
  document.getElementById('scale-value').textContent = Math.round(state.scale * 100) + '%'
  localStorage.setItem(SCALE_KEY, String(state.scale))
}

function updateBadge() {
  const badge = document.getElementById('cart-badge')
  if (state.cart.length > 0) {
    badge.textContent = state.cart.length
    badge.classList.remove('hidden')
  } else {
    badge.classList.add('hidden')
  }
}

function loadOrders() {
  try { return JSON.parse(localStorage.getItem(ORDERS_KEY) || '[]') } catch { return [] }
}

function saveOrder(order) {
  const orders = loadOrders()
  orders.push(order)
  localStorage.setItem(ORDERS_KEY, JSON.stringify(orders))
}

function generateOrderId() {
  const now = new Date()
  const d = now.toISOString().slice(0, 10).replace(/-/g, '')
  const t = now.toTimeString().slice(0, 8).replace(/:/g, '')
  return `${d}-${t}-${Math.floor(Math.random() * 900 + 100)}`
}

function exportOrders(orders) {
  const rows = orders.flatMap((order) =>
    order.items.map((item, i) => ({
      Дата: order.date, Время: order.time, '№ заказа': order.id,
      Категория: item.categoryName, Номенклатура: item.name,
      'Кол-во': item.quantity, 'Ед.': item.unit,
      'Цена за ед.': item.unitPrice, 'Сумма по строке': item.lineTotal,
      'Итого по заказу': i === 0 ? order.total : '',
    }))
  )
  const byDay = {}
  orders.forEach((o) => {
    if (!byDay[o.date]) byDay[o.date] = { count: 0, total: 0 }
    byDay[o.date].count++
    byDay[o.date].total += o.total
  })
  const summary = Object.entries(byDay).sort(([a], [b]) => a.localeCompare(b)).map(([date, s]) => ({
    Дата: date, 'Кол-во заказов': s.count, 'Сумма за день': s.total,
  }))
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Заказы')
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'По дням')
  XLSX.writeFile(wb, `заказы_${new Date().toISOString().slice(0, 10)}.xlsx`)
}

function renderCatalog() {
  const main = document.getElementById('main')
  let html = '<div class="catalog"><p class="catalog__hint">Раскройте категорию, выберите товар и укажите количество</p>'

  categories.forEach((cat) => {
    const isOpen = state.expanded[cat.id]
    html += `<section class="category-panel">
      <button type="button" class="category-panel__header" data-toggle="${cat.id}">
        <span>${cat.name}</span><span class="category-panel__chevron">${isOpen ? '▲' : '▼'}</span>
      </button>`

    if (isOpen) {
      const sel = state.selected[cat.id] || ''
      const qty = state.quantity[cat.id] || 1
      const item = getItem(sel)
      const unitPrice = item ? getPrice(item, qty) : 0

      html += `<div class="category-panel__body">
        <label class="field"><span class="field__label">Выберите номенклатуру</span>
          <select class="field__select" data-select="${cat.id}">
            <option value="">— не выбрано —</option>
            ${cat.itemIds.map((id) => {
              const it = getItem(id)
              return `<option value="${id}" ${sel === id ? 'selected' : ''}>${it.name}</option>`
            }).join('')}
          </select>
        </label>`

      if (item) {
        html += `<div class="item-detail">
          <label class="field"><span class="field__label">Количество (${item.unit})</span>
            <input type="number" class="field__input" min="1" value="${qty}" data-qty="${cat.id}" />
          </label>
          <div class="item-detail__prices">
            <div class="price-row"><span>Цена за ед.</span><strong>${formatPrice(unitPrice)}</strong></div>
            <div class="price-row price-row--total"><span>Сумма</span><strong>${formatPrice(unitPrice * qty)}</strong></div>
          </div>
          <button type="button" class="btn btn-primary btn-block" data-add="${cat.id}">В корзину</button>
        </div>`
      }
      html += '</div>'
    }
    html += '</section>'
  })
  html += '</div>'
  main.innerHTML = html
}

function renderCart() {
  const total = state.cart.reduce((s, i) => s + i.lineTotal, 0)
  const main = document.getElementById('main')

  if (state.cart.length === 0) {
    main.innerHTML = `<div class="cart">
      <button type="button" class="btn btn-ghost back-btn" id="back-btn">← Назад к каталогу</button>
      <div class="empty-state"><p>Корзина пуста</p>
        <button type="button" class="btn btn-primary" id="back-btn2">Перейти к выбору</button>
      </div></div>`
    return
  }

  main.innerHTML = `<div class="cart">
    <button type="button" class="btn btn-ghost back-btn" id="back-btn">← Назад к каталогу</button>
    <ul class="cart-list">
      ${state.cart.map((item) => `<li class="cart-item">
        <div><div class="cart-item__name">${item.name}</div>
          <div class="cart-item__meta">${item.quantity} ${item.unit} × ${formatPrice(item.unitPrice)}</div></div>
        <div class="cart-item__actions">
          <strong>${formatPrice(item.lineTotal)}</strong>
          <button type="button" class="btn btn-ghost btn-icon" data-remove="${item.id}" aria-label="Удалить">✕</button>
        </div></li>`).join('')}
    </ul>
    <div class="cart-total"><span>Итого</span><strong>${formatPrice(total)}</strong></div>
    <button type="button" class="btn btn-primary btn-block btn-lg" id="confirm-btn">Подтвердить заказ</button>
  </div>`
}

function render() {
  updateBadge()
  if (state.view === 'catalog') renderCatalog()
  else renderCart()
}

document.getElementById('main').addEventListener('click', (e) => {
  const t = e.target.closest('[data-toggle]')
  if (t) { state.expanded[t.dataset.toggle] = !state.expanded[t.dataset.toggle]; render(); return }

  const add = e.target.closest('[data-add]')
  if (add) {
    const catId = add.dataset.add
    const cat = categories.find((c) => c.id === catId)
    const item = getItem(state.selected[catId])
    const qty = state.quantity[catId] || 1
    if (!item || !cat) return
    const unitPrice = getPrice(item, qty)
    state.cart.push({
      id: `${item.id}-${Date.now()}`, nomenclatureId: item.id, name: item.name,
      categoryId: cat.id, categoryName: cat.name, unit: item.unit,
      quantity: qty, unitPrice, lineTotal: unitPrice * qty,
    })
    state.selected[catId] = ''
    state.quantity[catId] = 1
    state.expanded[catId] = false
    showToast(`«${item.name}» добавлен в корзину`)
    render()
    return
  }

  if (e.target.id === 'back-btn' || e.target.id === 'back-btn2') { state.view = 'catalog'; render(); return }

  const rem = e.target.closest('[data-remove]')
  if (rem) { state.cart = state.cart.filter((i) => i.id !== rem.dataset.remove); render(); return }

  if (e.target.id === 'confirm-btn') {
    const now = new Date()
    const order = {
      id: generateOrderId(),
      date: now.toLocaleDateString('ru-RU'),
      time: now.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' }),
      items: state.cart,
      total: state.cart.reduce((s, i) => s + i.lineTotal, 0),
    }
    saveOrder(order)
    exportOrders(loadOrders())
    state.cart = []
    state.view = 'catalog'
    showToast(`Заказ ${order.id} оформлен. Excel загружен.`)
    render()
  }
})

document.getElementById('main').addEventListener('change', (e) => {
  if (e.target.dataset.select) {
    state.selected[e.target.dataset.select] = e.target.value
    state.quantity[e.target.dataset.select] = 1
    render()
  }
})

document.getElementById('main').addEventListener('input', (e) => {
  if (e.target.dataset.qty) {
    state.quantity[e.target.dataset.qty] = Math.max(1, parseInt(e.target.value, 10) || 1)
    render()
  }
})

document.getElementById('cart-btn').addEventListener('click', () => { state.view = 'cart'; render() })
document.getElementById('export-btn').addEventListener('click', () => {
  const orders = loadOrders()
  if (!orders.length) { showToast('Нет сохранённых заказов'); return }
  exportOrders(orders)
  showToast('Все заказы выгружены в Excel')
})
document.getElementById('scale-down').addEventListener('click', () => {
  state.scale = Math.max(1, +(state.scale - 0.1).toFixed(1)); applyScale()
})
document.getElementById('scale-up').addEventListener('click', () => {
  state.scale = Math.min(1.8, +(state.scale + 0.1).toFixed(1)); applyScale()
})

applyScale()
render()
