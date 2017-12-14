import System
import Rhino
import Rhino.UI
import Eto.Drawing as drawing
import Eto.Forms as forms
import scriptcontext as sc
import rhinoscriptsyntax as rs

import sys
sys.path.append(r'C:\Users\Mostapha\Documents\code\ladybug-tools\ladybug')

try:
    import ladybug
except ImportError:
    raise ImportError('Failed to import ladybug. Install ladybug and try again!')

from ladybug.epw import EPW
from ladybug.sunpath import Sunpath
import ladybug.geometry as geo
from ladybug.location import Location

ladybug.isplus = True


class DrawSunPath(Rhino.Display.DisplayConduit):

    def __init__(self, curves, suns=None, color=None):
        self.curves = curves
        self.suns = suns or []
        self.color = color or System.Drawing.Color.Black

    # DrawOverlay override
    def DrawOverlay(self, e):
        for curve in self.curves:
            e.Display.DrawCurve(curve.ToNurbsCurve(), self.color, 2)
        for sun in self.suns:
            e.Display.DrawSphere(
                Rhino.Geometry.Sphere(sun, 5),
                System.Drawing.Color.Yellow,
                2)


class SunPathForm(forms.Form):

    # initializer
    def __init__(self):
        # place holder for epw file
        self.epw = None
        self.conduit = None

        # Basic form initialization
        self.initialize()

        # Create the form's controls
        self.create_form_controls()

        # Fill the form's listbox
        self.update_location_data()

        # Create Rhino event handlers
        self.create_events()

    # Basic form initialization
    def initialize(self):
        self.Title = 'Sunpath'
        self.Icon = drawing.Icon(r"asset/ladybug.ico")
        self.Padding = drawing.Padding(5)
        self.Resizable = True
        self.Maximizable = False
        self.Minimizable = False
        self.ShowInTaskbar = False
        self.MinimumSize = drawing.Size(200, 50)
        # FormClosed event handler
        self.Closed += self.OnFormClosed

    def update_location_data(self):
        """update location data."""
        if not self.epw:
            return
        # clear current items
        self.m_listbox.Items.Clear()
        item = forms.ListItem()
        item.Text = str(self.epw.location)
        self.m_listbox.Items.Add(item)

    # CloseDocument event handler
    def OnCloseDocument(self, sender, e):
        self.m_listbox.Items.Clear()

    # NewDocument event handler
    def OnNewDocument(self, sender, e):
        self.update_location_data()

    # EndOpenDocument event handler
    def OnEndOpenDocument(self, sender, e):
        self.update_location_data()

    # Create Rhino event handlers
    def create_events(self):
        Rhino.RhinoDoc.CloseDocument += self.OnCloseDocument
        Rhino.RhinoDoc.NewDocument += self.OnNewDocument
        Rhino.RhinoDoc.EndOpenDocument += self.OnEndOpenDocument

    # Remove Rhino event handlers
    def remove_events(self):
        Rhino.RhinoDoc.CloseDocument -= self.OnCloseDocument
        Rhino.RhinoDoc.NewDocument -= self.OnNewDocument
        Rhino.RhinoDoc.EndOpenDocument -= self.OnEndOpenDocument

    def create_form_controls(self):
        """Create all of the controls used by the form."""
        # Create table layout
        layout = forms.TableLayout()
        layout.Padding = drawing.Padding(10)
        layout.Spacing = drawing.Size(5, 5)
        # Add controls to layout
        layout.Rows.Add(forms.Label(Text='Location:'))
        layout.Rows.Add(self.create_list_box())
        layout.Rows.Add(self.create_button_row())
        # Set the content
        self.Content = layout

    def create_list_box(self):
        """
        Create the table row that contains the listbox.
        Called by create_form_controls
        """
        # Create the listbox
        self.m_listbox = forms.ListBox()
        self.m_listbox.Size = drawing.Size(200, 100)
        # Create the table row
        table_row = forms.TableRow()
        table_row.ScaleHeight = True
        table_row.Cells.Add(self.m_listbox)
        return table_row

    def on_load_epw(self, sender, e):
        """Import epw event handler."""
        filter = "EPW file (*.epw)|*.epw|All Files (*.*)|*.*||"
        epw_file = rs.OpenFileName("Open .epw Weather File", filter)
        # update location data based on epw file
        self.epw = EPW(epw_file)
        self.update_location_data()
        self.draw_sunpath()

    def draw_sunpath(self):
        """Calculate and draw sun path to Rhino."""
        self.clear_conduit()

        _location = self.epw.location
        north_ = 0
        _centerPt_ = None
        _scale_ = 1
        _sunScale_ = 1
        _annual_ = True

        daylight_saving_period = None  # temporary until we fully implement it
        _hoys_ = range(23)

        # initiate sunpath based on location
        sp = Sunpath.from_location(_location, north_, daylight_saving_period)

        # draw sunpath geometry
        sunpath_geo = \
            sp.draw_sunpath(_hoys_, _centerPt_, _scale_, _sunScale_, _annual_)

        analemma = sunpath_geo.analemma_curves
        compass = sunpath_geo.compass_curves
        daily = sunpath_geo.daily_curves

        sun_pts = sunpath_geo.sun_geos

        suns = sunpath_geo.suns
        vectors = (geo.vector(*sun.sun_vector) for sun in suns)
        altitudes = (sun.altitude for sun in suns)
        azimuths = (sun.azimuth for sun in suns)
        center_pt = _centerPt_ or geo.point(0, 0, 0)
        hoys = (sun.hoy for sun in suns)
        datetimes = (sun.datetime for sun in suns)

        # draw the curves to canvas
        self.conduit = DrawSunPath(list(analemma) + list(compass) + list(daily), sun_pts)
        self.conduit.Enabled = True
        # add lights
        for sun in suns:
            light = Rhino.Geometry.Light.CreateSunLight(
                north_, sun.azimuth - 90, sun.altitude)
            sc.doc.Lights.Add(light)
        sc.doc.Views.Redraw()

    def create_button_row(self):
        """
        Creates the table row that contains the button controls.
        Called by create_form_controls.
        """
        # Select button
        select_button = forms.Button(Text='Open EPW')
        select_button.Click += self.on_load_epw
        # Create layout
        layout = forms.TableLayout(Spacing=drawing.Size(5, 5))
        layout.Rows.Add(forms.TableRow(None, select_button, None))
        return layout

    def clear_conduit(self):
        """clear conduit."""
        if not self.conduit:
            return

        try:
            # for light in Rhino.RhinoDoc.ActiveDoc.Lights:
            for count in range(Rhino.RhinoDoc.ActiveDoc.Lights.Count):
                light = Rhino.RhinoDoc.ActiveDoc.Lights[count]
                if not light.IsDeleted:
                    # try to delete the light
                    Rhino.RhinoDoc.ActiveDoc.Lights.Delete(count, True)
        except Exception as e:
            print('Failed to remove lights!:\n{}'.format(e))
        finally:
            self.conduit.Enabled = False
            sc.doc.Views.Redraw()

    def OnFormClosed(self, sender, e):
        """Form Closed event handler."""
        # Remove the events added in the initializer
        self.remove_events()
        self.clear_conduit()
        # Dispose of the form and remove it from the sticky dictionary
        if 'ladybug_sunpath' in sc.sticky:
            form = sc.sticky['ladybug_sunpath']
            if form:
                form.Dispose()
                form = None
            sc.sticky.Remove('ladybug_sunpath')


def init_sunpath_form():

    # See if the form is already visible
    if 'ladybug_sunpath' in sc.sticky:
        return

    # Create and show form
    form = SunPathForm()
    form.Owner = Rhino.UI.RhinoEtoApp.MainWindow
    form.Show()
    # Add the form to the sticky dictionary so it
    # survives when the main function ends.
    sc.sticky['ladybug_sunpath'] = form


if __name__ == '__main__':
    try:
        init_sunpath_form()
    except Exception as e:
        print "Exception: {}".format(e)
