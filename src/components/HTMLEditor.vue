<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { QuillEditor } from '@vueup/vue-quill'
import '@vueup/vue-quill/dist/vue-quill.snow.css'

const props = defineProps({
  html: {
    type: String,
    default: ''
  }
})

const editor = ref()
const htmlString = ref('<p class="ql-align-center">Добрый день!</p><p class="ql-align-center">Уважаемый житель квартиры, располагающейся по адресу:</p><p class="ql-align-center">{{ streetLocal }}, кв. {{ apartmentLocal }}</p><p><br></p><p>По сведениям, предоставленным МФЦ района {{ districtLocal }} на {{ dateLocal }} г. у Вас имеется задолженность за жилищно-коммунальные услуги в размере {{ debtLocal.<span style="color: rgb(179, 146, 240);">toFixed</span>(<span style="color: rgb(121, 184, 255);">2</span>) }} руб., которую необходимо СРОЧНО ПОГАСИТЬ.</p><p><strong><u>Оплату необходимо производить по долговому Единому платежному документу (ЕПД) сформированному по Вашему коду плательщика № {{ payerCodeLocal }}.</u></strong></p><p>Документы, подтверждающие оплату, необходимо направить на электронную почту {{ emailLocal }}, либо предоставить в ГБУ «Жилищник района {{ districtLocal }}» по адресу:</p><p>г. Москва, {{ ukAddressLocal }}, контактный телефон {{ ukPhoneLocal }}. Получить долговой ЕПД Вы можете на информационном портале https:www.mos.ru/ или обратившись непосредственно в управляющую организацию</p><p><strong>ВНИМАНИЕ! Сумма задолженности в уведомлении может отличаться от задолженности, указанной в долговом ЕПД. Оплату необходимо производить только по долговому ЕПД</strong>. С уважением,</p><p>ГБУ «Жилищник района {{ districtLocal }}»<strong>{{ akvilaEmailLocal }}</strong></p>')

const initEditor = () => {
  // editor.value = new Quill('#editor', {
  //   modules: {
  //     toolbar: [
  //       ['bold', 'italic', 'underline'],
  //       ['link', 'image'], ['align'],
  //       [{ 'list': 'ordered' }, { 'list': 'bullet' }, { 'list': 'check' }],
  //       [{ 'size': ['small', false, 'large', 'huge'] }],  // custom dropdown
  //       [{ 'header': [1, 2, 3, 4, 5, 6, false] }]
  //     ],
  //   },
  //   placeholder: 'test',
  //   theme: 'snow'
  // })
  // editor.value.setText(props.html)
  // editor.value.clipboard.dangerouslyPasteHTML(0, props.html)
  // editor.value.getSelection()
  // const delta = editor.value.clipboard.convert({ html: props.html })
  // editor.value.setContents(delta)
}

onMounted(() => {
  initEditor()
})
</script>

<template>
  <QuillEditor
    ref="editor"
    theme="snow"
    toolbar="essential"
    content-type="html"
    :content="htmlString"
  />
</template>

<style scoped>

</style>
