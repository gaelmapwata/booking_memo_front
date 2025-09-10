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
</script> 