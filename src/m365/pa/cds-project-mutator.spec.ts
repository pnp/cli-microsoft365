import * as assert from 'assert';
import * as path from 'path';
import { XMLSerializer } from 'xmldom';
import CdsProjectMutator from './cds-project-mutator';

describe('CdsProjectMutator', () => {
  const pcfProjectFilePath: string = path.join('../path/to/projectDirectory', 'project.pcfproj');

  it('Executes when no project reference already exists', () => {
    const cdsProjectMutator = new CdsProjectMutator(`<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <PowerAppsTargetsPath>$(MSBuildExtensionsPath)\Microsoft\VisualStudio\v$(VisualStudioVersion)\PowerApps</PowerAppsTargetsPath>  
  </PropertyGroup>
  
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props')"/>

  <PropertyGroup>
    <ProjectGuid>00cad86b-4f87-4bcc-958c-b89640a96c21</ProjectGuid>
    <TargetFrameworkVersion>v4.6.2</TargetFrameworkVersion>
    <!--Remove TargetFramework when this is available in 16.1-->
    <TargetFramework>net462</TargetFramework>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
  </PropertyGroup>

  <!-- Solution Packager overrides, un-comment to use: SolutionPackagerType (Managed, Unmanaged, Both)
  <PropertyGroup>
    <SolutionPackageType>Managed</SolutionPackageType>
  </PropertyGroup>
  -->
  
  <ItemGroup>
    <PackageReference Include="Microsoft.PowerApps.MSBuild.Solution" Version="0.*"/>
  </ItemGroup>
  
  <ItemGroup>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\.gitignore"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\bin\**"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\obj\**"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.cdsproj"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.cdsproj.user"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.sln"/>
  </ItemGroup>
  
  <ItemGroup>
    <None Include="$(MSBuildThisFileDirectory)\**" Exclude="@(ExcludeDirectories)"/>
  </ItemGroup>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets')"/>

</Project>`);

    assert.doesNotThrow(() => {
      cdsProjectMutator.addProjectReference(pcfProjectFilePath)
    });

    assert.strictEqual(new XMLSerializer().serializeToString(cdsProjectMutator.cdsProjectDocument), `<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <PowerAppsTargetsPath>$(MSBuildExtensionsPath)\Microsoft\VisualStudio\v$(VisualStudioVersion)\PowerApps</PowerAppsTargetsPath>  
  </PropertyGroup>
  
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props')"/>

  <PropertyGroup>
    <ProjectGuid>00cad86b-4f87-4bcc-958c-b89640a96c21</ProjectGuid>
    <TargetFrameworkVersion>v4.6.2</TargetFrameworkVersion>
    <!--Remove TargetFramework when this is available in 16.1-->
    <TargetFramework>net462</TargetFramework>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
  </PropertyGroup>

  <!-- Solution Packager overrides, un-comment to use: SolutionPackagerType (Managed, Unmanaged, Both)
  <PropertyGroup>
    <SolutionPackageType>Managed</SolutionPackageType>
  </PropertyGroup>
  -->
  
  <ItemGroup>
    <PackageReference Include="Microsoft.PowerApps.MSBuild.Solution" Version="0.*"/>
  </ItemGroup>
  
  <ItemGroup>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\.gitignore"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\bin\**"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\obj\**"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.cdsproj"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.cdsproj.user"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.sln"/>
  </ItemGroup>
  
  <ItemGroup>
    <None Include="$(MSBuildThisFileDirectory)\**" Exclude="@(ExcludeDirectories)"/>
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="${pcfProjectFilePath}"/>
  </ItemGroup>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets')"/>

</Project>`);
  });

  it('Executes when a project reference already exists', () => {
    const cdsProjectMutator = new CdsProjectMutator(`<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <PowerAppsTargetsPath>$(MSBuildExtensionsPath)\Microsoft\VisualStudio\v$(VisualStudioVersion)\PowerApps</PowerAppsTargetsPath>  
  </PropertyGroup>
  
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props')"/>

  <PropertyGroup>
    <ProjectGuid>00cad86b-4f87-4bcc-958c-b89640a96c21</ProjectGuid>
    <TargetFrameworkVersion>v4.6.2</TargetFrameworkVersion>
    <!--Remove TargetFramework when this is available in 16.1-->
    <TargetFramework>net462</TargetFramework>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
  </PropertyGroup>

  <!-- Solution Packager overrides, un-comment to use: SolutionPackagerType (Managed, Unmanaged, Both)
  <PropertyGroup>
    <SolutionPackageType>Managed</SolutionPackageType>
  </PropertyGroup>
  -->
  
  <ItemGroup>
    <PackageReference Include="Microsoft.PowerApps.MSBuild.Solution" Version="0.*"/>
  </ItemGroup>
  
  <ItemGroup>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\.gitignore"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\bin\**"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\obj\**"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.cdsproj"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.cdsproj.user"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.sln"/>
  </ItemGroup>
  
  <ItemGroup>
    <None Include="$(MSBuildThisFileDirectory)\**" Exclude="@(ExcludeDirectories)"/>
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\path\to\projectDirectory\project2.pcfproj"/>
  </ItemGroup>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets')"/>

</Project>`);

    assert.doesNotThrow(() => {
      cdsProjectMutator.addProjectReference(pcfProjectFilePath)
    });

    assert.strictEqual(new XMLSerializer().serializeToString(cdsProjectMutator.cdsProjectDocument), `<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <PowerAppsTargetsPath>$(MSBuildExtensionsPath)\Microsoft\VisualStudio\v$(VisualStudioVersion)\PowerApps</PowerAppsTargetsPath>  
  </PropertyGroup>
  
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props')"/>

  <PropertyGroup>
    <ProjectGuid>00cad86b-4f87-4bcc-958c-b89640a96c21</ProjectGuid>
    <TargetFrameworkVersion>v4.6.2</TargetFrameworkVersion>
    <!--Remove TargetFramework when this is available in 16.1-->
    <TargetFramework>net462</TargetFramework>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
  </PropertyGroup>

  <!-- Solution Packager overrides, un-comment to use: SolutionPackagerType (Managed, Unmanaged, Both)
  <PropertyGroup>
    <SolutionPackageType>Managed</SolutionPackageType>
  </PropertyGroup>
  -->
  
  <ItemGroup>
    <PackageReference Include="Microsoft.PowerApps.MSBuild.Solution" Version="0.*"/>
  </ItemGroup>
  
  <ItemGroup>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\.gitignore"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\bin\**"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\obj\**"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.cdsproj"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.cdsproj.user"/>
    <ExcludeDirectories Include="$(MSBuildThisFileDirectory)\*.sln"/>
  </ItemGroup>
  
  <ItemGroup>
    <None Include="$(MSBuildThisFileDirectory)\**" Exclude="@(ExcludeDirectories)"/>
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\path\to\projectDirectory\project2.pcfproj"/>
    <ProjectReference Include="${pcfProjectFilePath}"/>
  </ItemGroup>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets')"/>

</Project>`);
  });

  it('Executes when no ItemGroups exists', () => {
    const cdsProjectMutator = new CdsProjectMutator(`<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <PowerAppsTargetsPath>$(MSBuildExtensionsPath)\Microsoft\VisualStudio\v$(VisualStudioVersion)\PowerApps</PowerAppsTargetsPath>  
  </PropertyGroup>
  
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props')"/>

  <PropertyGroup>
    <ProjectGuid>00cad86b-4f87-4bcc-958c-b89640a96c21</ProjectGuid>
    <TargetFrameworkVersion>v4.6.2</TargetFrameworkVersion>
    <!--Remove TargetFramework when this is available in 16.1-->
    <TargetFramework>net462</TargetFramework>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
  </PropertyGroup>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets')"/>

</Project>`);

    assert.doesNotThrow(() => {
      cdsProjectMutator.addProjectReference(pcfProjectFilePath)
    });

    assert.strictEqual(new XMLSerializer().serializeToString(cdsProjectMutator.cdsProjectDocument), `<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <PowerAppsTargetsPath>$(MSBuildExtensionsPath)\Microsoft\VisualStudio\v$(VisualStudioVersion)\PowerApps</PowerAppsTargetsPath>  
  </PropertyGroup>
  
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props')"/>

  <PropertyGroup>
    <ProjectGuid>00cad86b-4f87-4bcc-958c-b89640a96c21</ProjectGuid>
    <TargetFrameworkVersion>v4.6.2</TargetFrameworkVersion>
    <!--Remove TargetFramework when this is available in 16.1-->
    <TargetFramework>net462</TargetFramework>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="${pcfProjectFilePath}"/>
  </ItemGroup>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets')"/>

</Project>`);
  });

  it('Executes when no PropertyGroups exists', () => {
    const cdsProjectMutator = new CdsProjectMutator(`<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">

  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props')"/>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets')"/>

</Project>`);

    assert.doesNotThrow(() => {
      cdsProjectMutator.addProjectReference(pcfProjectFilePath)
    });

    assert.strictEqual(new XMLSerializer().serializeToString(cdsProjectMutator.cdsProjectDocument), `<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">

  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props"/>

  <ItemGroup>
    <ProjectReference Include="${pcfProjectFilePath}"/>
  </ItemGroup>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.props')"/>

  <Import Project="$(MSBuildToolsPath)\Microsoft.Common.targets"/>
  <Import Project="$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets" Condition="Exists('$(PowerAppsTargetsPath)\Microsoft.PowerApps.VisualStudio.Solution.targets')"/>

</Project>`);
  });

  it('Executes when no ImportGroups exists', () => {
    const cdsProjectMutator = new CdsProjectMutator(`<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
</Project>`);

    assert.doesNotThrow(() => {
      cdsProjectMutator.addProjectReference(pcfProjectFilePath)
    });

    assert.strictEqual(new XMLSerializer().serializeToString(cdsProjectMutator.cdsProjectDocument), `<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" DefaultTargets="Build" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">

  <ItemGroup>
    <ProjectReference Include="${pcfProjectFilePath}"/>
  </ItemGroup>
</Project>`);
  });
});
