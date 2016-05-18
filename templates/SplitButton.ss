<div $getAttributesHTML("class") class="ss-tabset $extraClass">
	<ul>
	<% loop $Tabs %>
        <li class="<% if $First %>first<% else_if $Last %>last<% end_if %> $MiddleString $extraClass"><a href="#$id" id="tab-$id">$Title</a></li>
	<% end_loop %>
	</ul>

	<% loop $Tabs %>
		<% if $Tabs %>
			$FieldHolder
		<% else %>
			<div $AttributesHTML>
				<% loop $Fields %>
					$FieldHolder
				<% end_loop %>
			</div>
		<% end_if %>
	<% end_loop %>
</div>
