class ExcelCalculator {
	constructor() {
		this.scriptUrl =
			'https://cors-anywhere.herokuapp.com/https://script.google.com/macros/s/AKfycbwKGCC9vAClV8eI-ubL0kyiYPnuq5QfEE0YXlA_g0yAedFelv63pXdG_mUqRWyXOgrL/exec'

		this.currentPage = 'customer'
		this.init()
	}

	init() {
		this.setupNavigation()
		this.setupEvents()
		this.loadData()
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
		if (pageName === 'results') this.calculateWithExcel()
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

		const originalText = calculateBtn.textContent
		calculateBtn.textContent = '⏳ Расчет выполняется...'
		calculateBtn.disabled = true

		try {
			const excelData = this.collectExcelData()
			console.log('📤 Отправляем данные в Excel:', excelData)

			const response = await fetch(this.scriptUrl, {
				method: 'POST',
				body: JSON.stringify(excelData),
			})

			const text = await response.text()
			console.log('🔍 Ответ сервера (сырой):', text)

			let result
			try {
				result = JSON.parse(text)
			} catch (err) {
				console.warn('⚠️ Не удалось разобрать JSON:', err)
				alert('Ответ от сервера не JSON, смотри консоль.')
				return
			}

			console.log('📊 Результаты из Excel:', result)

			if (result.success) {

				const clean = val => {
					if (val === undefined || val === null || val === '' || val === '#VALUE!') return 0
					const num = parseFloat(String(val).replace(',', '.'))
					return isNaN(num) ? 0 : num
				}

				const sheetsKg = clean(result.sheetsKg)
				const circulation = clean(result.circulation)
				const total = clean(result.total)
				const vat = clean(result.vat)
				const final = clean(result.final)

				document.getElementById('sheets-kg').value = sheetsKg
				document.getElementById('circulation').value = circulation

				this.showResults(total, vat, final)

				this.showMessage(
					total === 0
						? '⚠️ Нет данных для расчета'
						: '✅ Расчет выполнен успешно',
					total === 0 ? 'warning' : 'success'
				)
			} else {
				throw new Error(result.error || 'Ошибка при расчёте')
			}
		} catch (error) {
			console.error('❌ Ошибка:', error)
			this.showMessage(`Ошибка: ${error.message}`, 'error')
		} finally {
			calculateBtn.textContent = originalText
			calculateBtn.disabled = false
		}
	}

	showResults(total, vat, final) {
		const format = n =>
			isNaN(n) || n === undefined ? '—' : `${parseFloat(n).toFixed(2)} ₽`

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
			box.style.transition = 'opacity 0.3s ease'
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

	setupEvents() {
		const calculateBtn = document.querySelector('[data-next="results"]')
		if (calculateBtn)
			calculateBtn.addEventListener('click', e => {
				e.preventDefault()
				this.calculateWithExcel()
			})

		const updateRatesBtn = document.getElementById('update-rates-btn')
		if (updateRatesBtn)
			updateRatesBtn.addEventListener('click', () => this.updateCurrencyRates())

		const clearBtn = document.getElementById('clear-btn')
		if (clearBtn) clearBtn.addEventListener('click', () => this.clearData())

		const newCalcBtn = document.getElementById('new-calculation-btn')
		if (newCalcBtn) newCalcBtn.addEventListener('click', () => this.clearData())
	}

	async updateCurrencyRates() {
		const button = document.getElementById('update-rates-btn')
		if (!button) return

		const originalText = button.textContent
		button.textContent = '⏳ Загружаем...'
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
		} catch (e) {
			console.error('Ошибка:', e)
			this.showMessage('❌ Не удалось обновить курсы валют', 'error')
		} finally {
			button.textContent = originalText
			button.disabled = false
		}
	}

	saveData() {
		try {
			const data = this.collectExcelData()
			localStorage.setItem('calculator-excel-data', JSON.stringify(data))
		} catch (e) {
			console.error('Ошибка сохранения:', e)
		}
	}

	loadData() {
		try {
			const saved = localStorage.getItem('calculator-excel-data')
			if (!saved) return
			const data = JSON.parse(saved)
			for (const key in data) {
				const el = document.getElementById(
					key.replace(/[A-Z]/g, m => '-' + m.toLowerCase())
				)
				if (el) {
					el.value = data[key]
				}
			}
		} catch (e) {
			console.error('Ошибка загрузки данных:', e)
		}
	}

	clearData() {
		if (!confirm('Очистить все данные?')) return
		document.querySelectorAll('input, textarea, select').forEach(el => {
			if (el.type === 'checkbox') el.checked = false
			else el.value = ''
		})
		this.showResults(0, 0, 0)
		localStorage.removeItem('calculator-excel-data')
		this.showPage('customer')
	}
}

document.addEventListener('DOMContentLoaded', () => {
	window.calculator = new ExcelCalculator()
	console.log('✅ Калькулятор с Excel запущен!')
})
