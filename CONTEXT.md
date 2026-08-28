# Category Sales Permissions

This context describes which partners may sell products in a hierarchical category catalog and how imported permission instructions are executed.

## Language

**Category Tree**:
The authoritative hierarchy of categories in which products may be sold.
_Avoid_: Taxonomy, catalog structure

**Partner**:
An organization whose right to sell products is evaluated for categories in the Category Tree.
_Avoid_: Seller, merchant

**Allowed Category Set**:
The set of Category leaves in which a Partner may sell products. It is cleared before that Partner's Permission Instructions are executed.
_Avoid_: Whitelist, permissions

**Permission Instruction**:
An Excel row that tells the loader to add or remove a Category subtree for a Partner. Instructions retain their source order within each Partner's partition.
_Avoid_: Permission rule, access, setting

**Instruction Batch**:
A consecutive group of one Partner's Permission Instructions with the same `allow yes` or `allow no` operation. An add batch adds the union of its expanded subtrees; a remove batch removes that union.
_Avoid_: Rule group

**Effective Permission**:
The resulting membership of a Category in a Partner's allowed set after executing Instruction Batches in source order.
_Avoid_: Rule value, source value

**Fully Restricted Branch**:
A Category branch for which a Partner is denied in every sellable leaf after all Instruction Batches are executed.
_Avoid_: Blocked category

**Partially Restricted Branch**:
A Category branch for which a Partner is allowed in at least one sellable leaf and denied in at least one other sellable leaf.
_Avoid_: Mixed category

**Orphaned Instruction**:
A Permission Instruction whose category identifier is absent from the Category Tree. It is an import-blocking error for the entire source file.
_Avoid_: Unknown category, missing category

**Inactive Category**:
A Category that remains in the Category Tree but is marked as no longer current.
_Avoid_: Orphaned category

**Redundant Instruction**:
A valid Permission Instruction whose subtree is already fully covered by another instruction in the same Instruction Batch. It does not change the batch result but may be reported as a warning.
_Avoid_: Duplicate, error
