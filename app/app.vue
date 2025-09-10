<template>
  <v-app>
    <v-main>
      <v-container class="py-8">
        <h1 class="text-h5 mb-4">Booking Memo</h1>
        <div class="mb-4 d-flex ga-2">
          <v-btn to="/" color="secondary" variant="text" prepend-icon="mdi-home">Accueil</v-btn>
          <v-btn to="/health" color="secondary" variant="text" prepend-icon="mdi-heart-pulse">Health</v-btn>
          <v-btn to="/outils" color="secondary" variant="text" prepend-icon="mdi-tools">Outils</v-btn>
        </div>
        <NuxtPage />
      </v-container>
    </v-main>
  </v-app>
</template>

<script setup lang="ts">
</script>
