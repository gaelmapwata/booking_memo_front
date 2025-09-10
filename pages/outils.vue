<template>
  <v-container class="py-8">
    <h2 class="text-h5 mb-6">Outils documents</h2>

    <v-row>
      <v-col cols="12" md="6">
        <v-card class="mb-6" variant="elevated">
          <v-card-title class="text-subtitle-1">Écrire dans Excel</v-card-title>
          <v-card-text>
            <v-tabs v-model="excelTab" density="compact" class="mb-3">
              <v-tab value="path">Fichier local</v-tab>
              <v-tab value="upload">Upload</v-tab>
            </v-tabs>
            <v-window v-model="excelTab">
              <v-window-item value="path">
                <v-form ref="excelFormRef" @submit.prevent="submitExcelPath">
                  <v-text-field v-model="excel.filePath" label="Chemin du fichier .xlsx" required placeholder="C:\\Users\\hp\\Documents\\test.xlsx" />
                  <v-text-field v-model="excel.sheetName" label="Nom de la feuille (optionnel)" placeholder="Feuil1" />
                  <div class="d-flex ga-2 mb-2">
                    <v-btn size="small" variant="tonal" @click="loadSheets" prepend-icon="mdi-file-tree">Lister feuilles</v-btn>
                    <v-select v-if="sheets.length" v-model="excel.sheetName" :items="sheets" label="Sélectionner la feuille" density="compact" style="max-width: 280px" />
                  </div>
                  <div class="d-flex ga-2">
                    <v-text-field v-model.number="excel.row" type="number" min="1" label="Ligne" />
                    <v-text-field v-model.number="excel.col" type="number" min="1" label="Colonne" />
                    <v-text-field v-model="excel.cell" label="Cellule A1" />
                    <v-btn size="small" variant="tonal" @click="syncA1FromRC">RC→A1</v-btn>
                    <v-btn size="small" variant="tonal" @click="syncRCFromA1">A1→RC</v-btn>
                  </div>
                  <v-text-field v-model="excel.value" label="Valeur à écrire" required />
                  <div class="d-flex ga-2 mt-2">
                    <v-btn type="submit" color="primary" :loading="excelLoading" prepend-icon="mdi-file-excel">Écrire</v-btn>
                    <v-btn variant="text" @click="resetExcel" prepend-icon="mdi-restore">Réinitialiser</v-btn>
                  </div>
                </v-form>
              </v-window-item>
              <v-window-item value="upload">
                <v-form @submit.prevent="submitExcelUpload">
                  <v-file-input v-model="excel.file" label="Fichier .xlsx" accept=".xlsx" prepend-icon="mdi-upload" />
                  <v-text-field v-model="excel.sheetName" label="Nom de la feuille (optionnel)" placeholder="Feuil1" />
                  <div class="d-flex ga-2">
                    <v-text-field v-model.number="excel.row" type="number" min="1" label="Ligne" />
                    <v-text-field v-model.number="excel.col" type="number" min="1" label="Colonne" />
                    <v-text-field v-model="excel.cell" label="Cellule A1" />
                    <v-btn size="small" variant="tonal" @click="syncA1FromRC">RC→A1</v-btn>
                    <v-btn size="small" variant="tonal" @click="syncRCFromA1">A1→RC</v-btn>
                  </div>
                  <v-text-field v-model="excel.value" label="Valeur à écrire" required />
                  <div class="d-flex ga-2 mt-2">
                    <v-btn type="submit" color="primary" :loading="excelLoading" prepend-icon="mdi-file-upload">Uploader & Écrire</v-btn>
                    <v-btn variant="text" @click="resetExcel" prepend-icon="mdi-restore">Réinitialiser</v-btn>
                  </div>
                </v-form>
              </v-window-item>
            </v-window>
            <v-divider class="my-4" />
            <div class="text-subtitle-2 mb-2">Écriture par nom de zone</div>
            <div class="d-flex ga-2 mb-2">
              <v-text-field v-model="named.name" label="Nom (ou A1)" placeholder="ClientNom" />
              <v-text-field v-model="named.value" label="Valeur" />
              <v-btn size="small" color="primary" :loading="named.loading" @click="writeNamedPath" prepend-icon="mdi-pencil">Écrire (chemin)</v-btn>
              <v-btn size="small" color="primary" :loading="named.loading" @click="writeNamedUpload" prepend-icon="mdi-upload">Écrire (upload)</v-btn>
            </div>
            <v-divider class="my-4" />
            <div class="text-subtitle-2 mb-2">Prévisualisation</div>
            <div class="d-flex ga-2 mb-2">
              <v-text-field v-model.number="preview.maxRows" type="number" min="1" max="50" density="compact" label="Lignes" style="max-width: 120px" />
              <v-text-field v-model.number="preview.maxCols" type="number" min="1" max="50" density="compact" label="Colonnes" style="max-width: 140px" />
              <v-btn size="small" :loading="preview.loading" @click="doPreview" prepend-icon="mdi-eye">Voir</v-btn>
            </div>
            <div v-if="preview.data && preview.data.length">
              <v-table density="compact" class="border rounded">
                <thead>
                  <tr>
                    <th v-for="c in preview.data[0].length" :key="c">{{ colName(c) }}</th>
                  </tr>
                </thead>
                <tbody>
                  <tr v-for="(row, r) in preview.data" :key="r">
                    <td v-for="(cell, c) in row" :key="c">{{ cell }}</td>
                  </tr>
                </tbody>
              </v-table>
            </div>
          </v-card-text>
        </v-card>
      </v-col>

      <v-col cols="12" md="6">
        <v-card variant="elevated">
          <v-card-title class="text-subtitle-1">Remplacer dans Word</v-card-title>
          <v-card-text>
            <v-tabs v-model="wordTab" density="compact" class="mb-3">
              <v-tab value="path">Fichier local</v-tab>
              <v-tab value="upload">Upload</v-tab>
            </v-tabs>
            <v-window v-model="wordTab">
              <v-window-item value="path">
                <v-form ref="wordFormRef" @submit.prevent="submitWordPath">
                  <v-text-field v-model="word.templatePath" label="Chemin du modèle .docx" required />
                  <v-text-field v-model="word.outputPath" label="Chemin de sortie (optionnel)" />
                  <v-textarea v-model="word.replacementsText" label="Remplacements (JSON)" rows="6" placeholder='{"NOM":"Alice","DATE":"2025-09-10"}' />
                  <div class="d-flex ga-2 mt-2">
                    <v-btn type="submit" color="primary" :loading="wordLoading" prepend-icon="mdi-file-word">Remplacer</v-btn>
                    <v-btn variant="text" @click="resetWord" prepend-icon="mdi-restore">Réinitialiser</v-btn>
                  </div>
                </v-form>
              </v-window-item>
              <v-window-item value="upload">
                <v-form @submit.prevent="submitWordUpload">
                  <v-file-input v-model="word.file" label="Modèle .docx" accept=".docx" prepend-icon="mdi-upload" />
                  <v-text-field v-model="word.outputPath" label="Chemin de sortie (optionnel)" />
                  <v-textarea v-model="word.replacementsText" label="Remplacements (JSON)" rows="6" placeholder='{"NOM":"Alice","DATE":"2025-09-10"}' />
                  <div class="d-flex ga-2 mt-2">
                    <v-btn type="submit" color="primary" :loading="wordLoading" prepend-icon="mdi-file-upload">Uploader & Remplacer</v-btn>
                    <v-btn variant="text" @click="resetWord" prepend-icon="mdi-restore">Réinitialiser</v-btn>
                  </div>
                </v-form>
              </v-window-item>
            </v-window>
          </v-card-text>
        </v-card>
      </v-col>
    </v-row>

    <v-snackbar v-model="snackbar.show" :color="snackbar.color" :timeout="4000">
      {{ snackbar.message }}
    </v-snackbar>
  </v-container>
</template>

<style scoped>
.v-table tbody tr:nth-child(odd) { background: rgba(0,0,0,0.02); }
.v-table thead th { position: sticky; top: 0; background: white; z-index: 1; }
</style>

<script setup lang="ts">
const config = useRuntimeConfig()

const snackbar = reactive({ show: false, message: '', color: 'success' as 'success' | 'error' | 'info' })
function notify(message: string, color: 'success' | 'error' | 'info' = 'success') {
  snackbar.message = message
  snackbar.color = color
  snackbar.show = true
}

// Helpers: conversion RC <-> A1
function columnNumberToName(n: number) {
  let name = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    name = String.fromCharCode(65 + rem) + name
    n = Math.floor((n - 1) / 26)
  }
  return name
}
function columnNameToNumber(name: string) {
  let num = 0
  for (const ch of name.toUpperCase()) {
    if (ch < 'A' || ch > 'Z') return NaN
    num = num * 26 + (ch.charCodeAt(0) - 64)
  }
  return num
}

// Excel form state
const excelTab = ref<'path' | 'upload'>('path')
const excelFormRef = ref()
const excel = reactive<{ filePath: string; sheetName: string; cell: string; value: string; row: number | null; col: number | null; file: File | null }>({
  filePath: '',
  sheetName: '',
  cell: '',
  value: '',
  row: null,
  col: null,
  file: null
})
const excelLoading = ref(false)
function resetExcel() {
  excel.filePath = ''
  excel.sheetName = ''
  excel.cell = ''
  excel.value = ''
  excel.row = null
  excel.col = null
  excel.file = null
}
function syncA1FromRC() {
  if (excel.row && excel.col) {
    excel.cell = `${columnNumberToName(excel.col)}${excel.row}`
  }
}
function syncRCFromA1() {
  const m = /^([A-Za-z]+)(\d+)$/.exec(excel.cell || '')
  if (!m) return
  excel.col = columnNameToNumber(m[1])
  excel.row = Number(m[2])
}
async function submitExcelPath() {
  excelLoading.value = true
  try {
    const cell = excel.cell || (excel.row && excel.col ? `${columnNumberToName(excel.col)}${excel.row}` : '')
    if (!(excel.filePath && cell)) throw new Error('filePath et cell/A1 sont requis')
    const res = await $fetch(`${config.public.apiBase}/excel/write`, {
      method: 'POST',
      body: { filePath: excel.filePath, sheetName: excel.sheetName || undefined, cell, value: excel.value }
    })
    notify('Écriture Excel réussie')
    console.debug(res)
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    excelLoading.value = false
  }
}
async function submitExcelUpload() {
  excelLoading.value = true
  try {
    const cell = excel.cell || (excel.row && excel.col ? `${columnNumberToName(excel.col)}${excel.row}` : '')
    if (!(excel.file && cell)) throw new Error('fichier et cell/A1 sont requis')
    const form = new FormData()
    form.append('file', excel.file as any)
    form.append('sheetName', excel.sheetName || '')
    form.append('cell', cell)
    form.append('value', excel.value || '')
    const res = await $fetch(`${config.public.apiBase}/excel/write-upload`, {
      method: 'POST',
      body: form
    })
    notify('Upload & écriture Excel réussis')
    console.debug(res)
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    excelLoading.value = false
  }
}

// Word form state
const wordTab = ref<'path' | 'upload'>('path')
const wordFormRef = ref()
const word = reactive<{ templatePath: string; outputPath: string; replacementsText: string; file: File | null }>({
  templatePath: '',
  outputPath: '',
  replacementsText: '{"NOM":"Alice"}',
  file: null
})
const wordLoading = ref(false)
function resetWord() {
  word.templatePath = ''
  word.outputPath = ''
  word.replacementsText = '{"NOM":"Alice"}'
  word.file = null
}
async function submitWordPath() {
  wordLoading.value = true
  try {
    if (!word.templatePath) throw new Error('templatePath est requis')
    const replacements = word.replacementsText?.trim() ? JSON.parse(word.replacementsText) : {}
    const res = await $fetch(`${config.public.apiBase}/word/replace`, {
      method: 'POST',
      body: { templatePath: word.templatePath, outputPath: word.outputPath || undefined, replacements }
    })
    notify('Remplacement Word réussi')
    console.debug(res)
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    wordLoading.value = false
  }
}
async function submitWordUpload() {
  wordLoading.value = true
  try {
    if (!word.file) throw new Error('fichier requis')
    const form = new FormData()
    form.append('file', word.file as any)
    form.append('outputPath', word.outputPath || '')
    form.append('replacements', word.replacementsText || '{}')
    const res = await $fetch(`${config.public.apiBase}/word/replace-upload`, {
      method: 'POST',
      body: form
    })
    notify('Upload & remplacement Word réussis')
    console.debug(res)
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    wordLoading.value = false
  }
}

const preview = reactive<{ maxRows: number; maxCols: number; loading: boolean; data: any[] | null }>({ maxRows: 10, maxCols: 10, loading: false, data: null })
function colName(idx: number) { return columnNumberToName(idx) }
async function doPreview() {
  preview.loading = true
  preview.data = null
  try {
    if (excelTab.value === 'path') {
      if (!excel.filePath) throw new Error('filePath requis')
      const res: any = await $fetch(`${config.public.apiBase}/excel/preview`, { method: 'POST', body: { filePath: excel.filePath, sheetName: excel.sheetName || undefined, maxRows: preview.maxRows, maxCols: preview.maxCols } })
      preview.data = res?.data || []
    } else {
      if (!excel.file) throw new Error('fichier requis')
      const form = new FormData()
      form.append('file', excel.file as any)
      form.append('sheetName', excel.sheetName || '')
      form.append('maxRows', String(preview.maxRows))
      form.append('maxCols', String(preview.maxCols))
      const res: any = await $fetch(`${config.public.apiBase}/excel/preview-upload`, { method: 'POST', body: form })
      preview.data = res?.data || []
    }
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    preview.loading = false
  }
}

const sheets = ref<string[]>([])
async function loadSheets() {
  try {
    if (excelTab.value === 'path') {
      if (!excel.filePath) throw new Error('filePath requis')
      const res: any = await $fetch(`${config.public.apiBase}/excel/sheets`, { method: 'POST', body: { filePath: excel.filePath } })
      sheets.value = res?.sheets || []
    } else {
      if (!excel.file) throw new Error('fichier requis')
      const form = new FormData()
      form.append('file', excel.file as any)
      const res: any = await $fetch(`${config.public.apiBase}/excel/sheets-upload`, { method: 'POST', body: form })
      sheets.value = res?.sheets || []
    }
    if (sheets.value.length && !excel.sheetName) excel.sheetName = sheets.value[0]
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  }
}

const named = reactive<{ name: string; value: string; loading: boolean }>({ name: '', value: '', loading: false })
async function writeNamedPath() {
  named.loading = true
  try {
    if (!(excel.filePath && named.name)) throw new Error('filePath et name requis')
    const res = await $fetch(`${config.public.apiBase}/excel/write-named`, { method: 'POST', body: { filePath: excel.filePath, name: named.name, value: named.value } })
    notify('Écriture par nom réussie')
    console.debug(res)
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    named.loading = false
  }
}
async function writeNamedUpload() {
  named.loading = true
  try {
    if (!(excel.file && named.name)) throw new Error('fichier et name requis')
    const form = new FormData()
    form.append('file', excel.file as any)
    form.append('name', named.name)
    form.append('value', named.value)
    const res = await $fetch(`${config.public.apiBase}/excel/write-named-upload`, { method: 'POST', body: form })
    notify('Upload & écriture par nom réussis')
    console.debug(res)
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    named.loading = false
  }
}
</script> 