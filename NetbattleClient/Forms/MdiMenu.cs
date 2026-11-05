namespace Netbattle.Forms {
    internal static class MdiMenu {
        /// <summary>
        /// Sets up an MDI child's menu to merge into the container's. Menus the container already has (File, Help) take the
        /// child's items, the rest go in front of Help. The child's own strip is hidden and the room it took is given back.
        /// Call after InitializeComponent, with MdiParent already set.
        /// </summary>
        public static void MergeIntoParent(Form child, MenuStrip menu) {
            MenuStrip? parentMenu = child.MdiParent?.MainMenuStrip;
            if (parentMenu == null)
                return;

            List<string> parentMenus = parentMenu.Items.Cast<ToolStripItem>().Select(a => a.Text).ToList();
            int insertAt = parentMenus.IndexOf("&Help");
            if (insertAt < 0)
                insertAt = parentMenus.Count;

            foreach (ToolStripItem item in menu.Items) {
                if (parentMenus.Contains(item.Text)) {
                    item.MergeAction = MergeAction.MatchOnly;
                    continue;
                }

                item.MergeAction = MergeAction.Insert;
                item.MergeIndex = insertAt++;
            }

            int height = menu.Height;
            menu.Visible = false;

            // -- Docked controls take the space back themselves; move the rest up. Anchors are dropped while the form
            // -- shrinks so bottom-anchored controls don't move twice.
            List<Control> placed = child.Controls.Cast<Control>().Where(a => a != menu && a.Dock == DockStyle.None).ToList();
            Dictionary<Control, AnchorStyles> anchors = placed.ToDictionary(a => a, a => a.Anchor);

            child.SuspendLayout();
            foreach (Control control in placed) {
                control.Anchor = AnchorStyles.Top | AnchorStyles.Left;
                control.Top -= height;
            }

            child.ClientSize = new Size(child.ClientSize.Width, child.ClientSize.Height - height);

            foreach (Control control in placed)
                control.Anchor = anchors[control];
            child.ResumeLayout();
        }
    }
}
