<script setup>
import { createBlank, mount } from 'bymoffice';
import { onBeforeUnmount, onMounted, ref } from 'vue';

/** Shows `file` (an address) in the editor, or a new file of `kind` ('docx' or 'xlsx') when there is none. */
const props = defineProps({ file: String, kind: { type: String, default: 'xlsx' } });
const host = ref(null);
const error = ref('');

// Not reactive: the editor draws into the container itself, Vue never looks at it.
let view = null;
let gone = false;

onMounted(async () => {
  try {
    const mounted = await mount(host.value, {
      file: props.file || (await createBlank(props.kind)),
      // Where vite.config.js put the font folder of the package: next to the page, under the base of the app.
      fontsUrl: import.meta.env.BASE_URL + 'bymoffice-fonts/',
    });
    // mount() is asynchronous: the component can be gone before it resolves.
    if (gone) return mounted.destroy();
    view = mounted;
    window.bymEditor = mounted; // to try it out in the browser console
  } catch (e) {
    if (!gone) error.value = e instanceof Error ? e.message : String(e);
  }
});

onBeforeUnmount(() => {
  gone = true;
  if (view) view.destroy();
  if (window.bymEditor === view) window.bymEditor = undefined;
});
</script>

<template>
  <p v-if="error" class="error" role="alert">{{ error }}</p>
  <div ref="host" class="editor"></div>
</template>
