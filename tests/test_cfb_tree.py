"""
Validate the CFB directory red-black tree against the MS-CFB invariants that
the Windows Structured Storage implementation (and therefore Outlook/MAPI)
enforces. Lenient readers like olefile ignore these, so this is our best
Windows-compatibility guarantee short of running Outlook.
"""

import pytest

from pymsgkit import MSGWriter, create_email, RecipientType
from pymsgkit.cfb import CFBWriter, DirectoryEntry, EntryType, Color

NIL = DirectoryEntry.NOSTREAM


def _key(entries, did):
    name = entries[did].name[:31]
    return (len(name), name.upper())


def _validate_tree(entries, root_did):
    """Assert BST ordering + red-black properties for one sibling tree.

    Returns the in-order list of DIDs. Raises AssertionError on any violation.
    """
    # 1. In-order traversal must be sorted by the MS-CFB name comparison, and
    #    the tree must be acyclic / finite.
    order = []
    seen = set()

    def walk(d):
        if d == NIL:
            return
        assert d not in seen, "cycle in directory tree"
        seen.add(d)
        walk(entries[d].left_sibling)
        order.append(d)
        walk(entries[d].right_sibling)

    walk(root_did)
    keys = [_key(entries, d) for d in order]
    assert keys == sorted(keys), "sibling tree is not BST-ordered by CFB rules"

    # 2. The root of a sibling tree must be black.
    assert entries[root_did].color == Color.BLACK, "tree root must be black"

    # 3. Red nodes have black children; every root-to-leaf path has the same
    #    number of black nodes (equal black-height).
    def black_height(d):
        if d == NIL:
            return 1
        l, r = entries[d].left_sibling, entries[d].right_sibling
        if entries[d].color == Color.RED:
            if l != NIL:
                assert entries[l].color == Color.BLACK, "red node has red child"
            if r != NIL:
                assert entries[r].color == Color.BLACK, "red node has red child"
        lh, rh = black_height(l), black_height(r)
        assert lh == rh, "unequal black-height (invalid red-black tree)"
        return lh + (1 if entries[d].color == Color.BLACK else 0)

    black_height(root_did)
    return order


def _validate_all(cfb):
    entries = cfb.directory_entries
    for did, e in enumerate(entries):
        if e.entry_type in (EntryType.ROOT, EntryType.STORAGE) and e.child != NIL:
            _validate_tree(entries, e.child)


def test_tree_valid_simple_cfb():
    cfb = CFBWriter()
    for i in range(1):
        cfb.add_stream("only", b"x")
    cfb._finalize_directory_tree()
    _validate_all(cfb)


@pytest.mark.parametrize("n", [2, 3, 5, 8, 13, 21, 50, 128])
def test_tree_valid_many_siblings(n):
    cfb = CFBWriter()
    for i in range(n):
        cfb.add_stream(f"stream_{i:04d}", b"data")
    cfb._finalize_directory_tree()
    _validate_all(cfb)


def test_tree_valid_nested_storages():
    cfb = CFBWriter()
    for s in range(3):
        did = cfb.add_storage(f"storage_{s}")
        for i in range(7):
            cfb.add_stream(f"inner_{i}", b"data", parent_did=did)
    cfb._finalize_directory_tree()
    _validate_all(cfb)


def test_tree_valid_for_real_message(tmp_path):
    """A fully-populated MSG (message props, 2 recipients, 2 attachments,
    named-property storage) must produce valid trees at every level."""
    msg = MSGWriter()
    msg.set_subject("Tree Validation")
    msg.set_sender("a@b.com", "Alice")
    msg.set_body("<p>hi</p>", is_html=True)
    msg.add_recipient("to@x.com", "To Person", RecipientType.TO)
    msg.add_recipient("cc@x.com", "Cc Person", RecipientType.CC)
    msg.add_attachment("a.txt", b"one")
    msg.add_attachment("b.bin", bytes(5000))  # forces a regular-sector stream
    msg.save(str(tmp_path / "m.msg"))

    _validate_all(msg.cfb)


def test_names_sorted_shorter_first():
    """CFB compares by length first: 'z' must sort before 'aa'."""
    cfb = CFBWriter()
    cfb.add_stream("aa", b"x")
    cfb.add_stream("z", b"x")
    cfb._finalize_directory_tree()
    order = _validate_tree(cfb.directory_entries, cfb.directory_entries[0].child)
    names = [cfb.directory_entries[d].name for d in order]
    assert names == ["z", "aa"]
