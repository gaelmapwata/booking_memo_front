<template>
  <v-container class="py-8">
    <h2 class="text-h5 mb-6">Outils documents</h2>

    <v-row>
      <v-col cols="12" md="6">
        <v-card class="mb-6" variant="elevated">
          <v-card-title class="text-subtitle-1">Écrire dans Excel</v-card-title>
          <v-card-text>
            <v-form ref="excelFormRef" @submit.prevent="submitExcel">
              <v-text-field v-model="excel.filePath" label="Chemin du fichier .xlsx" required placeholder="C:\\Users\\hp\\Documents\\test.xlsx" />
              <v-text-field v-model="excel.sheetName" label="Nom de la feuille (optionnel)" placeholder="Feuil1" />
              <v-text-field v-model="excel.cell" label="Cellule (ex: B3)" required />
              <v-text-field v-model="excel.value" label="Valeur à écrire" required />
              <div class="d-flex ga-2 mt-2">
                <v-btn type="submit" color="primary" :loading="excelLoading" prepend-icon="mdi-file-excel">
                  Écrire
                </v-btn>
                <v-btn variant="text" @click="resetExcel" prepend-icon="mdi-restore">Réinitialiser</v-btn>
              </div>
            </v-form>
          </v-card-text>
        </v-card>
      </v-col>

      <v-col cols="12" md="6">
        <v-card variant="elevated">
          <v-card-title class="text-subtitle-1">Remplacer dans Word</v-card-title>
          <v-card-text>
            <v-form ref="wordFormRef" @submit.prevent="submitWord">
              <v-text-field v-model="word.templatePath" label="Chemin du modèle .docx" required placeholder="C:\\Users\\hp\\Documents\\modele.docx" />
              <v-text-field v-model="word.outputPath" label="Chemin de sortie (optionnel)" placeholder="C:\\Users\\hp\\Documents\\sortie.docx" />
              <v-textarea v-model="word.replacementsText" label="Remplacements (JSON)" rows="6" placeholder='{"NOM":"Alice","DATE":"2025-09-10"}' />
              <div class="d-flex ga-2 mt-2">
                <v-btn type="submit" color="primary" :loading="wordLoading" prepend-icon="mdi-file-word">
                  Remplacer
                </v-btn>
                <v-btn variant="text" @click="resetWord" prepend-icon="mdi-restore">Réinitialiser</v-btn>
              </div>
            </v-form>
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

// Excel form state
const excelFormRef = ref()
const excel = reactive({ filePath: '', sheetName: '', cell: '', value: '' })
const excelLoading = ref(false)
function resetExcel() {
  excel.filePath = ''
  excel.sheetName = ''
  excel.cell = ''
  excel.value = ''
}
async function submitExcel() {
  excelLoading.value = true
  try {
    if (!excel.filePath || !excel.cell) throw new Error('filePath et cell sont requis')
    const res = await $fetch(`${config.public.apiBase}/excel/write`, {
      method: 'POST',
      body: {
        filePath: excel.filePath,
        sheetName: excel.sheetName || undefined,
        cell: excel.cell,
        value: excel.value
      }
    })
    notify('Écriture Excel réussie')
    console.debug(res)
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    excelLoading.value = false
  }
}

// Word form state
const wordFormRef = ref()
const word = reactive({ templatePath: '', outputPath: '', replacementsText: '{"NOM":"Alice"}' })
const wordLoading = ref(false)
function resetWord() {
  word.templatePath = ''
  word.outputPath = ''
  word.replacementsText = '{"NOM":"Alice"}'
}
async function submitWord() {
  wordLoading.value = true
  try {
    if (!word.templatePath) throw new Error('templatePath est requis')
    let replacements: any = {}
    if (word.replacementsText?.trim()) {
      replacements = JSON.parse(word.replacementsText)
    }
    const res = await $fetch(`${config.public.apiBase}/word/replace`, {
      method: 'POST',
      body: {
        templatePath: word.templatePath,
        outputPath: word.outputPath || undefined,
        replacements
      }
    })
    notify('Remplacement Word réussi')
    console.debug(res)
  } catch (e: any) {
    notify(e?.data?.error || e?.message || 'Erreur', 'error')
  } finally {
    wordLoading.value = false
  }
}
</script> 