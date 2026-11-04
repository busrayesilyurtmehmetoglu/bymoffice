<script setup>
import { ref } from 'vue';
import OfficeEditor from './OfficeEditor.vue';

// What to open: ?file=<address of a docx, xlsx, doc, xls or RTF file>, or a new file (?new=docx, default xlsx).
const params = new URLSearchParams(location.search);
const file = params.get('file');
const kind = params.get('new') === 'docx' ? 'docx' : 'xlsx';
const shown = ref(true);
</script>

<template>
  <header>
    <button type="button" @click="shown = !shown">{{ shown ? 'Close editor' : 'Open editor' }}</button>
  </header>
  <OfficeEditor v-if="shown" :file="file" :kind="kind" />
</template>
