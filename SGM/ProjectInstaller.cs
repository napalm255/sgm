using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;

namespace SGM
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void serviceProcessInstaller_AfterInstall(object sender, InstallEventArgs e)
        {

        }

        private void serviceInstaller_AfterInstall(object sender, InstallEventArgs e)
        {

        }
    }
}