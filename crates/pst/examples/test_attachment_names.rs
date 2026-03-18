use outlook_pst::{
    ltp::prop_context::PropertyValue,
    messaging::{
        attachment::{AttachmentData},
        folder::Folder,
        store::Store,
    },
    ndb::node_id::NodeId,
};
use std::{fs, rc::Rc};

fn main() -> anyhow::Result<()> {
    let path = std::env::args()
        .nth(1)
        .expect("Usage: test_attachment_names <pst-file>");

    // Use dyn Store (same as Python would)
    let store = outlook_pst::open_store(&path)?;

    println!("Store codepage: {}", store.properties().codepage());

    let entry_id = store.properties().ipm_sub_tree_entry_id()?;
    let root_folder = store.open_folder(&entry_id)?;

    visit_folder(&store, &root_folder, 0)?;
    Ok(())
}

fn visit_folder(
    store: &Rc<dyn Store>,
    folder: &Rc<dyn Folder>,
    depth: usize,
) -> anyhow::Result<()> {
    let indent = "  ".repeat(depth);
    let name = folder.properties().display_name().unwrap_or_default();
    println!("{indent}[Folder] {name}");

    if let Some(contents) = folder.contents_table() {
        for row in contents.rows_matrix() {
            let ctx = contents.context();
            let cols = row.columns(ctx)?;

            let mut msg_nid = None;
            for (col, val) in ctx.columns().iter().zip(cols.iter()) {
                if col.prop_id() == 0x67F2 {
                    if let Some(val) = val {
                        if let Ok(PropertyValue::Integer32(nid)) =
                            contents.read_column(val, col.prop_type())
                        {
                            msg_nid = Some(nid as u32);
                        }
                    }
                }
            }

            if let Some(nid) = msg_nid {
                let nid = NodeId::from(nid);
                let entry_id = store.properties().make_entry_id(nid)?;
                match store.open_message(&entry_id, None) {
                    Ok(message) => {
                        let subject = message
                            .properties()
                            .get(0x0037)
                            .map(|v| format!("{v:?}"))
                            .unwrap_or_default();
                        println!("{indent}  [Message] {subject}");

                        if let Some(att_table) = message.attachment_table() {
                            let att_ctx = att_table.context();
                            for att_row in att_table.rows_matrix() {
                                let att_cols = att_row.columns(att_ctx)?;

                                let mut att_nid = None;
                                for (col, val) in
                                    att_ctx.columns().iter().zip(att_cols.iter())
                                {
                                    if col.prop_id() == 0x67F2 {
                                        if let Some(val) = val {
                                            if let Ok(PropertyValue::Integer32(n)) =
                                                att_table.read_column(val, col.prop_type())
                                            {
                                                att_nid = Some(NodeId::from(n as u32));
                                            }
                                        }
                                    }
                                }

                                if let Some(sub_node) = att_nid {
                                    // Use Message::open_attachment (works with dyn Message)
                                    match message.open_attachment(sub_node, None) {
                                        Ok(attachment) => {
                                            let short_name = attachment
                                                .properties()
                                                .get(0x3704)
                                                .map(|v| format!("{v:?}"))
                                                .unwrap_or_default();
                                            let long_name = attachment
                                                .properties()
                                                .get(0x3707)
                                                .map(|v| format!("{v:?}"))
                                                .unwrap_or_default();
                                            let method = attachment
                                                .properties()
                                                .attachment_method()
                                                .unwrap_or(-1);
                                            let has_data = attachment.data().is_some();

                                            println!(
                                                "{indent}    [Attachment] long_name={long_name} short_name={short_name} method={method} has_data={has_data}"
                                            );

                                            if let Some(data) = attachment.data() {
                                                match data {
                                                    AttachmentData::Binary(bin) => {
                                                        let filename = attachment
                                                            .properties()
                                                            .get(0x3707)
                                                            .or_else(|| {
                                                                attachment.properties().get(0x3704)
                                                            })
                                                            .map(|v| match v {
                                                                PropertyValue::Unicode(u) => {
                                                                    u.to_string()
                                                                }
                                                                PropertyValue::String8(s) => {
                                                                    s.to_string()
                                                                }
                                                                _ => "attachment.bin".to_string(),
                                                            })
                                                            .unwrap_or_else(|| {
                                                                "attachment.bin".to_string()
                                                            });

                                                        let out_path = format!("/tmp/{filename}");
                                                        fs::write(&out_path, bin.buffer())?;
                                                        println!(
                                                            "{indent}      -> Wrote {} bytes to {out_path}",
                                                            bin.buffer().len()
                                                        );
                                                    }
                                                    AttachmentData::Message(_) => {
                                                        println!(
                                                            "{indent}      -> Embedded message"
                                                        );
                                                    }
                                                }
                                            }
                                        }
                                        Err(e) => {
                                            println!(
                                                "{indent}    [Attachment error] sub_node={sub_node:?}: {e}"
                                            );
                                        }
                                    }
                                }
                            }
                        }
                    }
                    Err(e) => {
                        println!("{indent}  [Message error] {e}");
                    }
                }
            }
        }
    }

    if let Some(hierarchy) = folder.hierarchy_table() {
        for row in hierarchy.rows_matrix() {
            let ctx = hierarchy.context();
            let cols = row.columns(ctx)?;

            for (col, val) in ctx.columns().iter().zip(cols.iter()) {
                if col.prop_id() == 0x67F2 {
                    if let Some(val) = val {
                        if let Ok(PropertyValue::Integer32(nid)) =
                            hierarchy.read_column(val, col.prop_type())
                        {
                            let nid = NodeId::from(nid as u32);
                            let entry_id = store.properties().make_entry_id(nid)?;
                            if let Ok(sub_folder) = store.open_folder(&entry_id) {
                                visit_folder(store, &sub_folder, depth + 1)?;
                            }
                        }
                    }
                }
            }
        }
    }

    Ok(())
}
