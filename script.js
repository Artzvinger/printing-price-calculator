class ExcelCalculator {
	constructor() {
		this.scriptUrl =
			'https://cors-anywhere.herokuapp.com/https://script.google.com/macros/s/AKfycbxvR3wQ5JH4M9fy7jz9yz9O5ZIpn3WDP7-CgSfxzW41iFZ_BXoBctj_O5YemJM8tx4t/exec'
		this.currentPage = 'customer'
		this.init()
	}

	init() {
		this.setupNavigation()
		this.setupEvents()
		this.loadData()
		this.enableAutoSave()
	}

	setupNavigation() {
		document
			.querySelectorAll('.nav-btn')
			.forEach(btn =>
				btn.addEventListener('click', e => this.showPage(e.target.dataset.page))
			)
		document
			.querySelectorAll('.next-btn')
			.forEach(btn =>
				btn.addEventListener('click', e => this.showPage(e.target.dataset.next))
			)
		document
			.querySelectorAll('.prev-btn')
			.forEach(btn =>
				btn.addEventListener('click', e => this.showPage(e.target.dataset.prev))
			)
		const startBtn = document.querySelector('.start-btn')
		if (startBtn)
			startBtn.addEventListener('click', () =>
				this.showPage(startBtn.dataset.page)
			)
		const newCalcBtn = document.getElementById('new-calculation-btn')
		if (newCalcBtn)
			newCalcBtn.addEventListener('click', () => this.showPage('customer'))
	}

	showPage(pageName) {
		document
			.querySelectorAll('.page')
			.forEach(p => p.classList.remove('active'))
		const target = document.getElementById(`${pageName}-page`)
		if (target) target.classList.add('active')
		document
			.querySelectorAll('.nav-btn')
			.forEach(b => b.classList.remove('active'))
		const activeBtn = document.querySelector(`[data-page="${pageName}"]`)
		if (activeBtn) activeBtn.classList.add('active')
		this.currentPage = pageName
	}

	setupEvents() {
		const calc = document.querySelector('[data-next="results"]')
		if (calc)
			calc.addEventListener('click', e => {
				e.preventDefault()
				this.calculateWithExcel()
			})
		const updateRates = document.getElementById('update-rates-btn')
		if (updateRates)
			updateRates.addEventListener('click', () => this.updateCurrencyRates())
		const clear = document.getElementById('clear-btn')
		if (clear) clear.addEventListener('click', () => this.clearData())
	}

	collectExcelData() {
		const get = id => document.getElementById(id)?.value || ''
		return {
			companyName: get('company-name'),
			companyAddress: get('company-address'),
			companyContacts: get('company-contacts'),
			productName: get('product-name'),
			quantity: get('quantity'),
			perSheet: get('per-sheet'),
			notes: get('notes'),
			materialName: get('material-name'),
			materialType: get('material-type') || 'paper',
			materialPrice: get('material-price'),
			materialCurrency: get('material-currency') || 'RUB',
			gramsPerM2: get('grams-per-m2'),
			printWidth: get('print-width'),
			printHeight: get('print-height'),
			purchaseWidth: get('purchase-width'),
			purchaseHeight: get('purchase-height'),
			formatSize: get('format-size') || 'А3',
			usdRate: get('usd-rate'),
			eurRate: get('eur-rate'),
			opMaterial: get('material-type'),
			opCutting: get('cutting-format'),
			opPrinting: get('print-type'),
			opLamination: get('lamination'),
			opUV: get('uv-varnish'),
			opCutting2: get('cutting'),
			opEmbossing1: get('embossing1'),
			opEmbossing2: get('embossing2'),
			opDieCutting: get('die-cutting'),
			opGluing: get('gluing'),
			opBinding: get('binding'),
			shippingDate: get('shipping-date'),
		}
	}

	async calculateWithExcel() {
		const calculateBtn = document.querySelector('[data-next="results"]')
		if (!calculateBtn) return
		const pw = parseFloat(document.getElementById('print-width')?.value || 0)
		const ph = parseFloat(document.getElementById('print-height')?.value || 0)
		const purW = parseFloat(
			document.getElementById('purchase-width')?.value || 0
		)
		const purH = parseFloat(
			document.getElementById('purchase-height')?.value || 0
		)
		if (pw > purW || ph > purH) {
			this.showAlert(
				'❌ Ошибка формата!',
				`Формат печати не может быть больше формата закупочного.<br><br>
				📏 Печать: <b>${pw}×${ph} мм</b><br>
				📐 Закупка: <b>${purW}×${purH} мм</b>`
			)
			return
		}
		const originalText = calculateBtn.textContent
		calculateBtn.textContent = '⏳ Расчет выполняется...'
		calculateBtn.disabled = true
		try {
			const excelData = this.collectExcelData()
			const response = await fetch(this.scriptUrl, {
				method: 'POST',
				body: JSON.stringify(excelData),
			})
			const text = await response.text()
			let result
			try {
				result = JSON.parse(text)
			} catch {
				this.showAlert('Ошибка', 'Сервер вернул некорректный JSON')
				return
			}
			if (result.success) {
				const clean = val => {
					if (!val || val === '#VALUE!') return 0
					const num = parseFloat(String(val).replace(',', '.'))
					return isNaN(num) ? 0 : num
				}
				this.setValueSafe('sheets-kg', clean(result.sheetsKg))
				this.setValueSafe('circulation', clean(result.circulation))
				this.showResults(
					clean(result.total),
					clean(result.vat),
					clean(result.final)
				)
				this.showMessage('✅ Расчет выполнен успешно', 'success')
				this.showPage('results')
			} else {
				this.showAlert('Ошибка', result.error || 'Ошибка при расчете')
			}
		} catch (err) {
			this.showAlert('Ошибка', err.message)
		} finally {
			calculateBtn.textContent = originalText
			calculateBtn.disabled = false
		}
	}

	async updateCurrencyRates() {
		const button = document.getElementById('update-rates-btn')
		if (!button) return
		const originalText = button.textContent
		button.textContent = '⏳ Загрузка...'
		button.disabled = true
		try {
			const r = await fetch('https://www.cbr-xml-daily.ru/daily_json.js')
			if (!r.ok) throw new Error('Ошибка загрузки ЦБ РФ')
			const data = await r.json()
			document.getElementById('usd-rate').value =
				data.Valute.USD.Value.toFixed(2)
			document.getElementById('eur-rate').value =
				data.Valute.EUR.Value.toFixed(2)
			this.showMessage(
				`💱 Курсы обновлены: USD = ${data.Valute.USD.Value.toFixed(
					2
				)} ₽ | EUR = ${data.Valute.EUR.Value.toFixed(2)} ₽`,
				'success'
			)
			this.saveData()
		} catch (e) {
			this.showMessage('❌ Не удалось обновить курсы валют', 'error')
		} finally {
			button.textContent = originalText
			button.disabled = false
		}
	}

	enableAutoSave() {
		const elements = document.querySelectorAll('input, select, textarea')
		elements.forEach(el => {
			el.addEventListener('input', () => this.saveData())
			el.addEventListener('change', () => this.saveData())
		})
	}

	saveData() {
		try {
			const allData = {}
			document.querySelectorAll('input, select, textarea').forEach(el => {
				if (el.type === 'checkbox') allData[el.id] = el.checked
				else allData[el.id] = el.value
			})
			localStorage.setItem('calculator-excel-data', JSON.stringify(allData))
		} catch (e) {
			console.error('Ошибка при сохранении:', e)
		}
	}

	loadData() {
		try {
			const saved = localStorage.getItem('calculator-excel-data')
			if (!saved) return
			const data = JSON.parse(saved)
			Object.keys(data).forEach(id => {
				const el = document.getElementById(id)
				if (el) {
					if (el.type === 'checkbox') el.checked = data[id]
					else el.value = data[id]
				}
			})
			this.showMessage(
				'💾 Данные восстановлены из локального хранилища',
				'info'
			)
		} catch (e) {
			console.error('Ошибка при загрузке данных:', e)
		}
	}

	clearData() {
		if (!confirm('Очистить все данные?')) return
		document
			.querySelectorAll('input, textarea, select')
			.forEach(el =>
				el.type === 'checkbox' ? (el.checked = false) : (el.value = '')
			)
		this.showResults(0, 0, 0)
		localStorage.removeItem('calculator-excel-data')
		this.showPage('customer')
	}

	setValueSafe(id, value) {
		const el = document.getElementById(id)
		if (!el) return
		if ('value' in el) el.value = value
		else el.textContent = value
	}

	showResults(total, vat, final) {
		const format = n => (isNaN(n) ? '—' : `${parseFloat(n).toFixed(2)} ₽`)
		const set = (id, val) => {
			const el = document.getElementById(id)
			if (el) el.textContent = format(val)
		}
		set('total-result', total)
		set('vat-result', vat)
		set('final-result', final)
	}

	showMessage(text, type = 'info') {
		let box = document.getElementById('message-box')
		if (!box) {
			box = document.createElement('div')
			box.id = 'message-box'
			box.style.position = 'fixed'
			box.style.bottom = '20px'
			box.style.right = '20px'
			box.style.padding = '12px 20px'
			box.style.borderRadius = '10px'
			box.style.color = 'white'
			box.style.fontSize = '16px'
			box.style.fontWeight = '500'
			box.style.zIndex = '9999'
			box.style.transition = 'opacity 0.3s'
			document.body.appendChild(box)
		}
		const colors = {
			success: '#28a745',
			error: '#dc3545',
			warning: '#ffc107',
			info: '#007bff',
		}
		box.textContent = text
		box.style.background = colors[type] || colors.info
		box.style.opacity = '1'
		setTimeout(() => (box.style.opacity = '0'), 4000)
	}

	showAlert(title, message) {
		let alertBox = document.getElementById('center-alert')
		if (!alertBox) {
			alertBox = document.createElement('div')
			alertBox.id = 'center-alert'
			alertBox.style.position = 'fixed'
			alertBox.style.top = '50%'
			alertBox.style.left = '50%'
			alertBox.style.transform = 'translate(-50%, -50%) scale(0.9)'
			alertBox.style.background = 'white'
			alertBox.style.borderRadius = '15px'
			alertBox.style.boxShadow = '0 8px 30px rgba(0,0,0,0.25)'
			alertBox.style.padding = '30px 40px'
			alertBox.style.textAlign = 'center'
			alertBox.style.zIndex = '99999'
			alertBox.style.fontFamily = 'Inter, sans-serif'
			alertBox.style.transition = 'all 0.3s ease'
			alertBox.style.opacity = '0'
			const titleEl = document.createElement('h3')
			titleEl.id = 'alert-title'
			titleEl.style.marginBottom = '10px'
			titleEl.style.fontSize = '20px'
			titleEl.style.color = '#dc3545'
			titleEl.style.fontWeight = '700'
			const messageEl = document.createElement('div')
			messageEl.id = 'alert-message'
			messageEl.style.fontSize = '16px'
			messageEl.style.color = '#2d3748'
			const closeBtn = document.createElement('button')
			closeBtn.textContent = 'ОК'
			closeBtn.style.marginTop = '20px'
			closeBtn.style.padding = '10px 25px'
			closeBtn.style.border = 'none'
			closeBtn.style.borderRadius = '8px'
			closeBtn.style.background = '#dc3545'
			closeBtn.style.color = 'white'
			closeBtn.style.fontWeight = '600'
			closeBtn.style.cursor = 'pointer'
			closeBtn.addEventListener('click', () => {
				alertBox.style.opacity = '0'
				alertBox.style.transform = 'translate(-50%, -50%) scale(0.9)'
				setTimeout(() => alertBox.remove(), 300)
			})
			alertBox.append(titleEl, messageEl, closeBtn)
			document.body.appendChild(alertBox)
		}
		document.getElementById('alert-title').innerHTML = title
		document.getElementById('alert-message').innerHTML = message
		setTimeout(() => {
			alertBox.style.opacity = '1'
			alertBox.style.transform = 'translate(-50%, -50%) scale(1)'
		}, 10)
	}
}

document.addEventListener('DOMContentLoaded', () => {
	window.calculator = new ExcelCalculator()
	console.log('✅ Калькулятор с Excel запущен!')
})
