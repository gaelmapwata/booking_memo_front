<template>
  <v-container class="py-8">
    <h2 class="text-h6 mb-4">Vérifier l’API</h2>
    <v-btn :loading="pending" color="primary" prepend-icon="mdi-heart-pulse" @click="check">
      Appeler /health
    </v-btn>
    <div class="mt-4">
      <div v-if="error" class="text-error">Erreur: {{ error?.message }}</div>
      <div v-else-if="data">Réponse: {{ data }}</div>
    </div>
  </v-container>
</template>

<script setup lang="ts">
const config = useRuntimeConfig()
const data = ref<string | null>(null)
const error = ref<any>(null)
const pending = ref(false)

async function check() {
  pending.value = true
  data.value = null
  error.value = null
  try {
    const res = await $fetch(`${config.public.apiBase}/health`)
    data.value = JSON.stringify(res)
  } catch (e: any) {
    error.value = e
  } finally {
    pending.value = false
  }
}
</script> 