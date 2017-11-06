package cl.majeanis.satelite.util;

import javax.annotation.PostConstruct;

import org.apache.commons.configuration2.Configuration;
import org.apache.commons.configuration2.FileBasedConfiguration;
import org.apache.commons.configuration2.PropertiesConfiguration;
import org.apache.commons.configuration2.builder.ConfigurationBuilderEvent;
import org.apache.commons.configuration2.builder.ReloadingFileBasedConfigurationBuilder;
import org.apache.commons.configuration2.builder.fluent.Parameters;
import org.apache.commons.configuration2.event.EventListener;
import org.springframework.stereotype.Component;

@Component
public class SisProperties
{
    private Configuration config;
    
    @PostConstruct
    private void init()
    {
        final ReloadingFileBasedConfigurationBuilder<FileBasedConfiguration> builder = 
                new ReloadingFileBasedConfigurationBuilder<FileBasedConfiguration>(PropertiesConfiguration.class)
                                .configure( new Parameters().properties()
                                    .setFileName("satelite.properties")
                                    .setThrowExceptionOnMissing(false)
                                    .setIncludesAllowed(true));
        builder.setAutoSave(true);
        builder.addEventListener(ConfigurationBuilderEvent.CONFIGURATION_REQUEST,
                new EventListener<ConfigurationBuilderEvent>()
                {
                    @Override
                    public void onEvent(ConfigurationBuilderEvent event)
                    {
                        builder.getReloadingController().checkForReloading(null);
                    }
                });
        
        try
        {
            config = builder.getConfiguration();
        } catch(Exception e)
        {
            config = null;
        }
    }

    public String getServerAD()
    {
        return config.getString("ad_server", "127.0.0.1");
    }
    
    public String getDomainAD()
    {
        return config.getString("ad_domain", "LOCAL_AD.CL");
    }
}
