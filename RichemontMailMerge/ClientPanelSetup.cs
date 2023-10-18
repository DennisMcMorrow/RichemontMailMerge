using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RichemontMailMerge
{
    public class ClientPanelSetup
{
    // Properties for panels and buttons
    public Panel Panel11 { get; set; }
    public Panel Panel12 { get; set; }
    // ... Add other panels

    public Button Button4 { get; set; }
    public Button Button5 { get; set; }
    // ... Add other buttons

    public void SetupForClient(string selectedClient)
    {
        switch (selectedClient)
        {
            case "Richemont":
                SetupForRichemont();
                break;
            case "Sunland":
                SetupForSunland();
                break;
            case "Primetals":
                // Handle Primetals setup if needed
                break;
            case "Caromont":
                // Handle Caromont setup if needed
                break;
            default:
                // Handle default case if needed
                break;
        }
    }

    private void SetupForRichemont()
    {
          
        }

    private void SetupForSunland()
    {
        Panel11.Visible = false;
        Panel12.Visible = true;
        // ... Set other panels and buttons for Sunland
    }

    // ... Add other private methods for other clients
}
}
