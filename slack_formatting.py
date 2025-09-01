"""Slack formatting helpers for consistent bot responses."""

import config


class SlackResponses:
    """Pre-formatted response templates for common bot interactions."""
    
    @staticmethod
    def error(message: str) -> str:
        """Format an error message."""
        return config.markdown_to_slack(f"❌ **Error:** {message}")
    
    @staticmethod
    def success(message: str) -> str:
        """Format a success message."""
        return config.markdown_to_slack(f"✅ {message}")
    
    @staticmethod
    def warning(message: str) -> str:
        """Format a warning message."""
        return config.markdown_to_slack(f"⚠️ {message}")
    
    @staticmethod
    def info(message: str) -> str:
        """Format an info message."""
        return config.markdown_to_slack(f"ℹ️ {message}")
    
    @staticmethod
    def proposal_confirmation(proposal_type: str, locations: list, client: str, details: dict) -> str:
        """Format a proposal confirmation message."""
        if proposal_type == "combined":
            message = f"📦 **Combined Package Proposal**\n\n"
            message += f"**Client:** {client}\n"
            message += f"**Locations:** {', '.join(locations)}\n"
            message += f"**Package Rate:** {details.get('combined_rate', 'Not specified')}\n\n"
            message += "📄 _Generating your proposal..._"
        else:
            message = f"📊 **Proposal{'s' if len(locations) > 1 else ''}**\n\n"
            message += f"**Client:** {client}\n"
            message += f"**Location{'s' if len(locations) > 1 else ''}:** {', '.join(locations)}\n\n"
            if details.get('durations'):
                message += f"**Duration Options:** {', '.join(details['durations'])}\n"
            message += "📄 _Generating your proposal{'s' if len(locations) > 1 else ''}..._"
        
        return config.markdown_to_slack(message)
    
    @staticmethod
    def location_list(locations: list) -> str:
        """Format a list of available locations."""
        if not locations:
            return config.markdown_to_slack("📍 No locations available. Use **'add location'** to add one.")
        
        message = "📍 **Available Locations:**\n\n"
        for loc in sorted(locations):
            message += f"• {loc}\n"
        
        return config.markdown_to_slack(message)
    
    @staticmethod
    def help_message() -> str:
        """Format the help message."""
        message = """🤖 **BackLite Media Proposal Bot**

I can help you create financial proposals for digital advertising locations.

**Available Commands:**
• Generate a proposal - Just describe what you need
• `list locations` - Show all available locations
• `add location` - Add a new location template (admin only)
• `refresh templates` - Reload location templates

**Examples:**
• _"Create a proposal for The Landmark, starting Jan 1st, 2 weeks at 1.5M"_
• _"I need proposals for landmark and gateway with different durations"_
• _"Combined package for jawhara, oryx and triple crown at 2 million total"_

**Tips:**
• For separate proposals, each location can have multiple duration/rate options
• For combined packages, specify one total rate for all locations
• Always include the client name for tracking"""
        
        return config.markdown_to_slack(message)
    
    @staticmethod
    def proposal_summary(result: dict) -> str:
        """Format a summary of generated proposals."""
        if result.get("is_combined"):
            message = f"✅ **Combined Package Generated Successfully**\n\n"
            message += f"📍 **Locations:** {result['locations']}\n"
            message += f"📄 **File:** {result['pdf_filename']}"
        elif result.get("is_single"):
            message = f"✅ **Proposal Generated Successfully**\n\n"
            message += f"📍 **Location:** {result['location']}\n"
            message += f"📄 **Files:** PowerPoint and PDF versions"
        else:
            message = f"✅ **Multiple Proposals Generated Successfully**\n\n"
            message += f"📍 **Locations:** {result['locations']}\n"
            message += f"📄 **Files:** Individual PowerPoints + Combined PDF"
        
        return config.markdown_to_slack(message)