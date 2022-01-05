package org.work.new_file_command_plugin;

import com.intellij.ide.actions.CreateFileAction;
import com.intellij.openapi.actionSystem.AnAction;
import com.intellij.openapi.actionSystem.AnActionEvent;
import com.intellij.openapi.actionSystem.CommonDataKeys;
import com.intellij.openapi.project.Project;
import com.intellij.openapi.ui.Messages;
import com.intellij.pom.Navigatable;
import org.jetbrains.annotations.NotNull;
import org.jetbrains.annotations.Nullable;

import javax.swing.*;


public class NewFileWorkAction extends CreateFileAction {

  /**
   * This default constructor is used by the IntelliJ Platform framework to instantiate this class based on plugin.xml
   * declarations. Only needed in {@link NewFileWorkAction} class because a second constructor is overridden.
   *
   * @see AnAction#AnAction()
   */
  public NewFileWorkAction() {
    super();
  }

  /**
   * This constructor is used to support dynamically added menu actions.
   * It sets the text, description to be displayed for the menu item.
   * Otherwise, the default AnAction constructor is used by the IntelliJ Platform.
   *
   * @param text        The text to be displayed as a menu item.
   * @param description The description of the menu item.
   * @param icon        The icon to be used with the menu item.
   */
  public NewFileWorkAction(@Nullable String text, @Nullable String description, @Nullable Icon icon) {
    super(text, description, icon);
  }

}
