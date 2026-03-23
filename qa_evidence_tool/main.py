from ttkthemes import ThemedTk
from views.ui_app import TestAssistantApp

if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    
    app = TestAssistantApp(root)
    
    root.mainloop()