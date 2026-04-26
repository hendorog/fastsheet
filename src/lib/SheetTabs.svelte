<script lang="ts">
  type Props = {
    sheetNames: string[];
    activeIndex: number;
    onSelect: (index: number) => void | Promise<void>;
    onAddSheet: () => void | Promise<void>;
    onTabContextMenu: (x: number, y: number, sheetIdx: number) => void;
  };

  let { sheetNames, activeIndex, onSelect, onAddSheet, onTabContextMenu }: Props = $props();
</script>

<div class="sheet-tabs" role="tablist">
  {#each sheetNames as name, i}
    <button
      type="button"
      class="sheet-tab"
      class:active={i === activeIndex}
      onclick={() => onSelect(i)}
      oncontextmenu={(e) => {
        e.preventDefault();
        onTabContextMenu(e.clientX, e.clientY, i);
      }}
    >
      {name}
    </button>
  {/each}
  <button
    type="button"
    class="sheet-tab add-tab"
    title="Add sheet"
    onclick={onAddSheet}
  >+</button>
</div>

<style>
  .sheet-tabs {
    display: flex;
    gap: 0.15rem;
    background: #e8e8e8;
    border-top: 1px solid #d0d0d0;
    padding: 0.15rem 0.4rem 0;
    overflow-x: auto;
    overflow-y: hidden;
    flex: 0 0 auto;
  }
  .sheet-tab {
    background: #f3f3f3;
    color: #444;
    border: 1px solid #c0c0c0;
    border-bottom: none;
    border-radius: 3px 3px 0 0;
    padding: 0.15rem 0.7rem;
    font: inherit;
    font-size: 11px;
    white-space: nowrap;
    cursor: pointer;
  }
  .sheet-tab:hover {
    background: #fff;
  }
  .sheet-tab.active {
    background: #fff;
    color: #1f6feb;
    border-color: #c0c0c0;
    font-weight: 700;
  }
  .sheet-tab.add-tab {
    color: #1f6feb;
    font-weight: 700;
    padding: 0.15rem 0.5rem;
  }
</style>
