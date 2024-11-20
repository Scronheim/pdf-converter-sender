<script setup lang="ts">
import { ref, onMounted, computed } from 'vue'
import html2pdf from 'html2pdf.js'
import * as XLSX from 'xlsx'
import { ElNotification } from 'element-plus'
import { Setting, Folder, Message, Delete } from '@element-plus/icons-vue'

import type { TableInstance } from 'element-plus'
import type { FileListItem } from '../types'


const data = ref<FileListItem[]>([])  //Реактивная переменная для хранения данных из Excel
const fileList = ref<string[]>([])

const streetLocal = ref('')
const apartmentLocal = ref('')
const districtLocal = ref('')
const dateLocal = ref('')
const debtLocal = ref(0)
const payerCodeLocal = ref('')
const emailLocal = ref('')
const ukAddressLocal = ref('')
const ukPhoneLocal = ref('')
const akvilaEmailLocal = ref('')

const user = ref({
  selectedPdfFolderPath: '',
  smtpHost: '',
  smtpLogin: '',
  smtpPassword: '',
})

const mailServers = [
  { title: 'Mail.ru', value: 'smtp.mail.ru' },
  { title: 'Yandex', value: 'smtp.yandex.ru' }
]
const settingsDialog = ref<boolean>(false)
const pdfTable = ref<TableInstance>()

const handleFile = (event): void => {
  const file = event.raw  //Получаем выбранный файл
  if (file) {
    const reader = new FileReader()  //Создаем новый FileReader
    reader.onload = async (e: ProgressEvent) => {
      const arrayBuffer = (e.target as FileReader).result as ArrayBuffer  //Получаем результат чтения файла
      const binaryStr = new Uint8Array(arrayBuffer).reduce((data, byte) => data + String.fromCharCode(byte), '')  //Конвертируем в бинарный строковый формат
      const workbook = XLSX.read(binaryStr, { type: 'binary' })  //Читаем файл как рабочую книгу
      const firstSheetName = workbook.SheetNames[0]  //Получаем имя первого листа
      const worksheet = workbook.Sheets[firstSheetName]  //Получаем первый лист
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })  //Конвертируем лист в JSON формат
      data.value = jsonData.slice(1)  //Сохраняем данные, начиная со второй строки (A2)
      await generatePdf()
    }
    reader.readAsArrayBuffer(file)  //Читаем файл как массив байтов
  } else {
    ElNotification({
      type: 'error',
      message: 'Произошла ошибка при чтении из Excel файла'
    })
  }
}

const generatePdf = async (): Promise<void> => {
  for (const [index, row] of data.value.entries()) {
    const [
      street,
      apartment,
      payerCode,
      akvilaEmail,
      debt,
      date,
      district,
      email,
      ukAddress,
      ukPhone
    ] = row
    
    streetLocal.value = street
    apartmentLocal.value = apartment
    payerCodeLocal.value = payerCode
    akvilaEmailLocal.value = akvilaEmail
    debtLocal.value = debt
    dateLocal.value = date
    districtLocal.value = district
    emailLocal.value = email
    ukAddressLocal.value = ukAddress
    ukPhoneLocal.value = ukPhone
    const blob = await html2pdf().from(document.getElementById('template')).output('blob')
    const reader = new FileReader()
    reader.onload = async () => {
      await window.ipcRenderer.invoke('downloadFile', reader.result, `${akvilaEmail}.pdf`)
      await getFileList()
      pdfTable.value?.toggleAllSelection()
    }
    reader.readAsArrayBuffer(blob)

    if (index === data.value.length - 1) {
      ElNotification({
        type: 'success',
        message: 'Файлы PDF сгенерированы'
      })
    }
  }
}

const openPdfFolderDialog = async (): Promise<void> => {
  const path = await window.ipcRenderer.invoke('selectPdfFolder')
  if (path) user.value.selectedPdfFolderPath = path
  await getFileList()
}

const loadUserSettings = async (): Promise<void> => {
  user.value = await window.ipcRenderer.invoke('getUserSettings')
}

const saveUserSettings = async (): Promise<void> => {
  await window.ipcRenderer.invoke('createMailTransport')
  await window.ipcRenderer.invoke('saveUserSettings', JSON.stringify(user.value))
  ElNotification({
    type: 'success',
    message: 'Настройки сохранены'
  })
  settingsDialog.value = false
}

const getFileList = async () => {
  fileList.value = await window.ipcRenderer.invoke('getPdfFileList')
}

const sendMail = async (row) => {
  await window.ipcRenderer.invoke('sendMail', row.email, 'Привет', `${user.value.selectedPdfFolderPath}/${row.name}`)
}

const sendAllSelectedMail = async () => {
  const rows = pdfTable.value?.getSelectionRows()
  if (rows.length) {
    rows.forEach(row => {
      sendMail(row)
    })
    ElNotification({
      type: 'success',
      message: 'Все выбранные письма отправлены'
    })
  } else {
    ElNotification({
      type: 'error',
      message: 'Нет выбранных файлов'
    })
  }
}

const removeFile = async (row) => {
  window.ipcRenderer.invoke('removeFile', `${user.value.selectedPdfFolderPath}/${row.name}`)
  await getFileList()
}

onMounted(async () => {
  await loadUserSettings()
  await window.ipcRenderer.invoke('createMailTransport')
  await getFileList()
})
</script>

<template>
  <div class="flex gap-3">
    <el-upload
      :limit="1"
      :show-file-list="false"
      :auto-upload="false"
      accept=".xls,.xlsx"
      v-model:file-list="data"
      @change="handleFile"
    >
      <el-button type="primary">
        Выберите Excel файл
      </el-button>
    </el-upload>
    <el-button :icon="Setting" @click="settingsDialog = true" />
  </div>

  <el-button
    v-if="fileList.length > 0"
    type="success"
    @click="sendAllSelectedMail"
  >
    Отправить выбранные
  </el-button>
  <el-table
    ref="pdfTable"
    :data="fileList"
    empty-text="PDF файлов не найдено"
  >
    <el-table-column type="selection" width="55" />
    <el-table-column
      fixed
      prop="name"
      label="Имя файла"
    />
    <el-table-column
      prop="actions"
      label="Действия"
    >
      <template #default="{row}">
        <el-button :icon="Message" @click="sendMail(row)">
          Отправить
        </el-button>
        <el-button
          type="danger"
          :icon="Delete"
          @click="removeFile(row)"
        />
      </template>
    </el-table-column>
  </el-table>
  <div style="display: none;">
    <div
      style="font-family: 'Times New Roman'; text-align: center; border: 1px solid #000; width: 210mm; height: 290mm; padding: 20px; box-sizing: border-box; color: black"
      id="template"
    >
      <p style="font-size: 22px; font-weight: bold;">
        <br>
        Добрый день!
        <br>
        Уважаемый житель квартиры, располагающейся по адресу:
        <br>
        {{ streetLocal }}, кв. {{ apartmentLocal }}
      </p>
      <br>
      <br>
      <p style="font-size: 18px;">
        По сведениям, предоставленным МФЦ района {{ districtLocal }} на {{ dateLocal }} г. у Вас имеется задолженность
        за
        жилищно-коммунальные услуги в размере {{ debtLocal.toFixed(2) }} руб., которую необходимо СРОЧНО ПОГАСИТЬ.
      </p>
      <br>
      <p style="font-size: 20px; font-weight: bold; text-decoration: underline;">
        Оплату необходимо производить по долговому Единому платежному документу (ЕПД) сформированному по Вашему коду
        плательщика № {{ payerCodeLocal }}.
      </p>
      <br>
      <p style="font-size: 18px;">
        Документы, подтверждающие оплату, необходимо направить на электронную почту {{ emailLocal }}, либо
        предоставить
        в ГБУ
        «Жилищник района {{ districtLocal }}» по адресу:
      </p>
      <br>

      <p style="font-size: 18px;">
        г. Москва, {{ ukAddressLocal }},
      </p>
      <br>

      <p style="font-size: 18px;">
        контактный телефон {{ ukPhoneLocal }}.
      </p>
      <br>

      <p style="font-size: 20px; font-weight: bold;">
        Получить долговой ЕПД Вы можете на информационном портале https:www.mos.ru/ или обратившись непосредственно
        в
        управляющую организацию.
      </p>
      <br>

      <p style="font-size: 20px; font-weight: bold;">
        ВНИМАНИЕ! Сумма задолженности в уведомлении может отличаться от задолженности, указанной в долговом ЕПД.
        Оплату
        необходимо производить только по долговому ЕПД.
      </p>
      <br>
      <br>
      <p style="text-align: left;font-size: 18px;">
        С уважением,
        <br>

        ГБУ «Жилищник района {{ districtLocal }}»
      </p>
      <br>
      <br>

      <p p style="text-align: left;font-size: 18px;font-weight: bold;">
        {{ akvilaEmailLocal }}
      </p>
    </div>
  </div>
   

  <el-dialog title="Настройки" v-model="settingsDialog">
    <div>
      <label for="selectedPdfFolderPath">Путь до папки с PDF</label>
      <el-input
        id="selectedPdfFolderPath"
        v-model="user.selectedPdfFolderPath"
        @click="openPdfFolderDialog"
      >
        <template #append>
          <el-button :icon="Folder" @click="openPdfFolderDialog" />
        </template>
      </el-input>
    </div>
    <div>
      <label for="smtpServer">Сервер почты</label>
      <el-select v-model="user.smtpHost">
        <el-option
          v-for="server in mailServers"
          :key="server.value"
          :label="server.title"
          :value="server.value"
        />
      </el-select>
    </div>
    <div>
      <label for="smtpLogin">Логин от почты</label>
      <el-input
        id="smtpLogin"
        v-model="user.smtpLogin"
      />
    </div>
    <div>
      <label for="smtpPassword">Пароль от почты</label>
      <el-input
        id="smtpPassword"
        v-model="user.smtpPassword"
      />
    </div>
    <template #footer>
      <el-button type="danger" @click="settingsDialog = false">
        Закрыть
      </el-button>
      <el-button type="success" @click="saveUserSettings">
        Сохранить
      </el-button>
    </template>
  </el-dialog>
</template>

<style lang="css">

</style>
